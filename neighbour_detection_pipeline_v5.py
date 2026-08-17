"""
Neighbour Detection Pipeline (v5)
==================================
Combines OS Places API (address/UPRN lookup) with OS NGD API (building AND
land parcel polygons) to identify TRUE boundary-touching neighbours for a
target property, then applies the client's business rules.

Layered defences against false-positive "neighbours":
  1. Buffered .intersects() as a coarse first pass (buildings + land parcels).
  2. shares_real_boundary(): the intersection must be an actual LINE of
     meaningful length, not just a single touching corner POINT.
  3. is_clean_two_party_boundary(): rejects a shared boundary if a THIRD
     property's parcel sits right at the MIDPOINT of that boundary — a
     signature of a multi-way "pinch point" where 3+ plots meet at a
     corner. Confirmed threshold: genuine neighbours had nearest-third-
     party distances of ~3.9-9m+; pinch-point false positives were ~1.5m.
     PINCH_POINT_CHECK_DISTANCE_M is set at 2.0m to sit in that gap.
  4. sanity_check_distance(): final straight-line distance check between
     the target's and candidate's actual address points. Catches cases
     the polygon logic can't — e.g. cul-de-sac ("Close") developments
     where two unrelated properties both border the same large shared/
     communal land parcel (a turning circle, verge, amenity strip), which
     makes them register as "touching" even though they are nowhere near
     each other in reality.

STATUS / STILL TUNING:
  MIN_THIRD_PARTY_AREA_SQM and MAX_PLAUSIBLE_NEIGHBOUR_DISTANCE_M are
  starting values based on the test addresses worked through so far, not
  exhaustively validated. Recommended before going live: build a table of
  15-20 known addresses (using client-style aerial screenshots to confirm
  expected neighbours) and run find_neighbours() on all of them after any
  threshold change, checking nothing that used to pass now fails and vice
  versa. Tune one constant at a time.

Requirements:
    pip install requests shapely

You will need an OS Data Hub API key on the Premium Plan (free tier /
development mode is fine for testing) with these APIs added to the project:
    - OS Places API
    - OS NGD API - Features
"""

import requests
from shapely.geometry import shape, Point

OS_API_KEY = "aJARAMAGjcqsJbfGyA9pXOAuYkwhyxHm"

PLACES_BASE = "https://api.os.uk/search/places/v1"
NGD_BASE = "https://api.os.uk/features/ngd/ofa/v1/collections"

BUFFER_M = 1.0                        # slack distance for polygon-polygon adjacency test
SEARCH_PAD_M = 30.0                   # bbox padding around the target property, in metres
ADDRESS_POINT_SLACK_M = 2.0           # buffer applied to a polygon before checking address points
ADDRESS_MATCH_FALLBACK_M = 6.0        # nearest-address-point fallback distance
LAND_CONTAINMENT_FALLBACK_M = 5.0     # nearest-land-parcel fallback distance (target's own parcel)
MIN_SHARED_BOUNDARY_LENGTH_M = 0.3    # minimum length to count as a real shared edge (not a corner)
PINCH_POINT_CHECK_DISTANCE_M = 2.0    # confirmed against real data — see notes above
MIN_THIRD_PARTY_AREA_SQM = 30.0       # ignore tiny driveway/path slivers in the pinch-point check
MAX_PLAUSIBLE_NEIGHBOUR_DISTANCE_M = 30.0  # final address-point sanity check — tune against test set
MAX_NEIGHBOURS_BEFORE_ABORT = 4

CRS_TEMPLATE = "http://www.opengis.net/def/crs/EPSG/0/{}"

# Classification codes that should be EXCLUDED even though they start with 'R'
# (flats, maisonettes, multi-dwelling buildings). Confirm/expand this list
# against the OS classification scheme document before going live.
EXCLUDED_RESIDENTIAL_SUBTYPES = {
    "RD06",  # Flat (converted)
    "RD07",  # Flat (Purpose Built)
    "RD08",  # Flat (Maisonette) -- confirm exact codes with OS scheme doc
}

srs = "EPSG:27700"  # set dynamically from the geocode response


# ---------------------------------------------------------------------------
# STEP 1 — Geocode the target address
# ---------------------------------------------------------------------------
def geocode_address(address_text):
    global srs
    url = f"{PLACES_BASE}/find"
    params = {"query": address_text, "maxresults": 1, "key": OS_API_KEY}
    r = requests.get(url, params=params)
    r.raise_for_status()
    data = r.json()
    srs = data.get("header", {}).get("output_srs", srs)
    results = data.get("results", [])
    

    if not results:
        return None
    return results[0]["DPA"]


def _crs():
    return CRS_TEMPLATE.format(srs.split(":")[-1])


# ---------------------------------------------------------------------------
# STEP 2 — Fetch polygons by TOID / by bbox, for a given NGD collection
# ---------------------------------------------------------------------------
def get_polygon_by_toid(collection, toid):
    url = f"{NGD_BASE}/{collection}/items"
    params = {
        "key": OS_API_KEY,
        "filter": f"toid='{toid}'",
        "filter-lang": "cql-text",
        "crs": _crs(),
    }
    r = requests.get(url, params=params)
    if r.status_code != 200:
        print(f"[{collection}] by-TOID error body:", r.text)
    r.raise_for_status()
    features = r.json().get("features", [])
    
    return shape(features[0]["geometry"]) if features else None


def get_nearby_polygons(collection, x, y, pad=SEARCH_PAD_M):
    bbox = f"{x - pad},{y - pad},{x + pad},{y + pad}"
    url = f"{NGD_BASE}/{collection}/items"
    params = {
        "key": OS_API_KEY,
        "bbox": bbox,
        "bbox-crs": _crs(),
        "crs": _crs(),
        "limit": 10,
    }
    r = requests.get(url, params=params)
    if r.status_code != 200:
        print(f"[{collection}] bbox error body:", r.text)
    r.raise_for_status()
    
    return r.json().get("features", [])


def find_containing_polygon(x, y, features, fallback_max_m=LAND_CONTAINMENT_FALLBACK_M):
    """Return (toid, polygon) of whichever feature's polygon contains point (x, y).
    Falls back to the nearest polygon within fallback_max_m if no exact
    containment is found (handles address points sitting on/near an edge)."""
    pt = Point(x, y)
    best_toid, best_poly, best_dist = None, None, None
    for feat in features:
        poly = shape(feat["geometry"])
        if poly.covers(pt):
            return feat["properties"]["toid"], poly
        d = poly.distance(pt)
        if best_dist is None or d < best_dist:
            best_toid, best_poly, best_dist = feat["properties"]["toid"], poly, d
    
    if best_dist is not None and best_dist < fallback_max_m:
        print(f"Warning: no exact containing parcel found, using nearest ({best_dist:.2f}m away)")
        return best_toid, best_poly
    return None, None


# ---------------------------------------------------------------------------
# STEP 3 — Real geometric adjacency test
# ---------------------------------------------------------------------------
def shares_real_boundary(poly_a, poly_b, min_shared_length=MIN_SHARED_BOUNDARY_LENGTH_M):
    """True only if the two polygons share an actual boundary LINE (not
    just a single touching corner point), of at least min_shared_length.
    Returns the shared geometry itself (for further checks) or None."""
    intersection = poly_a.boundary.intersection(poly_b.boundary)

    if intersection.is_empty:
        return None
    if intersection.geom_type in ("Point", "MultiPoint"):
        return None

    if hasattr(intersection, "geoms"):
        total_len = sum(g.length for g in intersection.geoms if hasattr(g, "length"))
    else:
        total_len = intersection.length

    if total_len < min_shared_length:
        return None
    
    return intersection


def is_clean_two_party_boundary(shared_segment, target_toid, candidate_toid,
                                  all_land_features,
                                  min_distance_from_third=PINCH_POINT_CHECK_DISTANCE_M,
                                  min_third_party_area=MIN_THIRD_PARTY_AREA_SQM):
    """Reject a shared boundary if a THIRD property's parcel sits right at
    its MIDPOINT — a signature of a multi-way pinch point (3+ plots meeting
    at a corner) rather than a genuine two-property boundary. Small parcels
    (driveway/path slivers) are ignored so they don't wrongly veto a real
    neighbour relationship."""
    midpoint = shared_segment.interpolate(0.5, normalized=True)
    for feat in all_land_features:
        toid = feat["properties"]["toid"]
        if toid in (target_toid, candidate_toid):
            continue
        poly = shape(feat["geometry"])
        if poly.area < min_third_party_area:
            continue
        if midpoint.distance(poly) < min_distance_from_third:
            return False
    
    return True


def find_touching_polygons(target_toid, target_polygon, nearby_features, all_land_features=None):
    buffered_target = target_polygon.buffer(BUFFER_M)
    touching = []
    for feat in nearby_features:
        toid = feat["properties"]["toid"]
        if toid == target_toid:
            continue
        poly = shape(feat["geometry"])

        if not buffered_target.intersects(poly):
            continue

        shared = shares_real_boundary(target_polygon, poly)
        if shared is None:
            continue

        if all_land_features is not None:
            if not is_clean_two_party_boundary(shared, target_toid, toid, all_land_features):
                continue

        touching.append({"toid": toid, "polygon": poly})
    
    return touching


# ---------------------------------------------------------------------------
# STEP 4 — Nearby address points (for matching polygons back to addresses)
# ---------------------------------------------------------------------------
def get_nearby_address_points(x, y, radius=SEARCH_PAD_M):
    url = f"{PLACES_BASE}/radius"
    params = {"point": f"{x},{y}", "radius": radius, "key": OS_API_KEY}
    r = requests.get(url, params=params)
    r.raise_for_status()
    
    return [res["DPA"] for res in r.json().get("results", []) if "DPA" in res]


def match_address_to_polygon(polygon, address_points,
                              buffer_m=ADDRESS_POINT_SLACK_M,
                              fallback_max_m=ADDRESS_MATCH_FALLBACK_M):
    """Find address point(s) belonging to this polygon. Tries a buffered
    .covers() test first, then falls back to the single nearest address
    point if it's genuinely close (handles edge-of-polygon placements)."""
    poly_buffered = polygon.buffer(buffer_m)
    matched = [dpa for dpa in address_points
               if poly_buffered.covers(Point(dpa["X_COORDINATE"], dpa["Y_COORDINATE"]))]
    if matched:
        return matched

    best_dpa, best_dist = None, None
    for dpa in address_points:
        pt = Point(dpa["X_COORDINATE"], dpa["Y_COORDINATE"])
        d = polygon.distance(pt)
        if best_dist is None or d < best_dist:
            best_dpa, best_dist = dpa, d

    if best_dist is not None and best_dist < fallback_max_m:
        return [best_dpa]
    
    return []


def sanity_check_distance(target_x, target_y, candidate_dpa,
                           max_dist=MAX_PLAUSIBLE_NEIGHBOUR_DISTANCE_M):
    """Final straight-line check between the target's and candidate's actual
    address points. Catches cases the polygon logic can't reliably detect —
    e.g. two properties both bordering a large shared/communal parcel
    (a cul-de-sac turning circle, a verge) that makes them register as
    'touching' even though they are nowhere near each other in reality."""
    cand_pt = Point(candidate_dpa["X_COORDINATE"], candidate_dpa["Y_COORDINATE"])
    target_pt = Point(target_x, target_y)
    d = target_pt.distance(cand_pt)
    
    return d <= max_dist, d


# ---------------------------------------------------------------------------
# STEP 5 — Client's filtering rules
# ---------------------------------------------------------------------------
def is_valid_neighbour(dpa_record):
    code = dpa_record.get("CLASSIFICATION_CODE", "")
    state = dpa_record.get("BLPU_STATE_CODE_DESCRIPTION", "")
    
    if not code.startswith("R"):
        return False, "not residential"
    if code in EXCLUDED_RESIDENTIAL_SUBTYPES:
        return False, "flat/maisonette excluded"
    if state and state.lower() != "in use":
        return False, "not an active/in-use property"

    return True, "ok"


# ---------------------------------------------------------------------------
# MAIN PIPELINE
# ---------------------------------------------------------------------------
def find_neighbours(address_text, debug=False):
    target = geocode_address(address_text)
    if not target:
        return {"error": "target address not found"}

    target_x, target_y = target["X_COORDINATE"], target["Y_COORDINATE"]
    target_building_toid = target.get("TOPOGRAPHY_LAYER_TOID")

    # ---- Layer 1: buildings ----
    target_building_poly = None
    if target_building_toid:
        target_building_poly = get_polygon_by_toid("bld-fts-buildingpart-1", target_building_toid)

    nearby_buildings = get_nearby_polygons("bld-fts-buildingpart-1", target_x, target_y)

    touching_building = []
    if target_building_poly is not None:
        touching_building = find_touching_polygons(
            target_building_toid, target_building_poly, nearby_buildings
        )

    # ---- Layer 2: land parcels (pinch-point veto applied) ----
    nearby_land = get_nearby_polygons("lnd-fts-land-1", target_x, target_y)
    target_land_toid, target_land_poly = find_containing_polygon(target_x, target_y, nearby_land)

    touching_land = []
    if target_land_poly is not None:
        touching_land = find_touching_polygons(
            target_land_toid, target_land_poly, nearby_land,
            all_land_features=nearby_land
        )

    if debug:
        print("target_building_toid:", target_building_toid)
        print("target_building_poly is None?", target_building_poly is None)
        print("nearby_buildings count:", len(nearby_buildings))
        print("nearby_land count:", len(nearby_land))
        print("target_land_toid:", target_land_toid)
        print("touching_building count:", len(touching_building))
        print("touching_land count:", len(touching_land))

    all_touching = touching_building + touching_land

    address_points = get_nearby_address_points(target_x, target_y)

    neighbours = []
    rejected = []
    seen_uprns = set()

    for t in all_touching:
        matched = match_address_to_polygon(t["polygon"], address_points)

        if not matched:
            rejected.append({"toid": t["toid"], "reason": "no address point found inside polygon"})
            continue

        for dpa in matched:
            uprn = dpa["UPRN"]
            if uprn == target.get("UPRN") or uprn in seen_uprns:
                continue

            # Final sanity check — catches shared-communal-parcel false positives
            ok, dist = sanity_check_distance(target_x, target_y, dpa)
            if not ok:
                seen_uprns.add(uprn)
                rejected.append({
                    "toid": t["toid"], "uprn": uprn, "address": dpa.get("ADDRESS"),
                    "reason": f"address point {dist:.1f}m away — exceeds plausible neighbour distance"
                })
                continue

            seen_uprns.add(uprn)

            valid, reason = is_valid_neighbour(dpa)
            if valid:
                neighbours.append({
                    "uprn": uprn,
                    "address": dpa["ADDRESS"],
                    "classification": dpa.get("CLASSIFICATION_CODE_DESCRIPTION"),
                    "toid": t["toid"],
                    "distance_m": round(dist, 1),
                })
            else:
                rejected.append({
                    "toid": t["toid"], "uprn": uprn,
                    "address": dpa.get("ADDRESS"), "reason": reason,
                })

    target_is_flat = target.get("CLASSIFICATION_CODE") in EXCLUDED_RESIDENTIAL_SUBTYPES
    aborted = target_is_flat and len(neighbours) > MAX_NEIGHBOURS_BEFORE_ABORT
    
    return {
        "target": {
            "uprn": target["UPRN"],
            "address": target["ADDRESS"],
            "building_toid": target_building_toid,
            "land_toid": target_land_toid,
            "classification": target.get("CLASSIFICATION_CODE_DESCRIPTION"),
        },
        "neighbours": neighbours,
        "rejected_candidates": rejected,
        "aborted_multi_dwelling": aborted,
        "neighbour_count": len(neighbours),
    }


# ---------------------------------------------------------------------------
# DEBUG HELPER — inspect WHY a boundary is being vetoed, before tuning
# MIN_THIRD_PARTY_AREA_SQM or PINCH_POINT_CHECK_DISTANCE_M.
# ---------------------------------------------------------------------------
def debug_clean_boundary_verbose(target_land_toid, candidate_toid, nearby_land_feats, top_n=5):
    target_poly = None
    cand_poly = None
    for feat in nearby_land_feats:
        if feat["properties"]["toid"] == target_land_toid:
            target_poly = shape(feat["geometry"])
        if feat["properties"]["toid"] == candidate_toid:
            cand_poly = shape(feat["geometry"])

    if target_poly is None or cand_poly is None:
        print("Could not find one or both polygons in nearby_land_feats")
        return

    shared = target_poly.boundary.intersection(cand_poly.boundary)
    if shared.is_empty or shared.geom_type in ("Point", "MultiPoint"):
        print("No real shared boundary between these two.")
        return

    midpoint = shared.interpolate(0.5, normalized=True)
    print(f"Shared boundary length: {shared.length:.2f}m")

    distances = []
    for feat in nearby_land_feats:
        toid = feat["properties"]["toid"]
        if toid in (target_land_toid, candidate_toid):
            continue
        poly = shape(feat["geometry"])
        d = midpoint.distance(poly)
        distances.append((d, toid, poly.area))
    
    distances.sort(key=lambda x: x[0])
    print(f"Nearest {top_n} other parcels to this boundary's midpoint:")
    for d, toid, area in distances[:top_n]:
        print(f"  {toid}: {d:.3f}m away, area={area:.1f} sqm")


if __name__ == "__main__":
    import json
    address = "9 Buxton Close Epsom KT19 8BD"
    result = find_neighbours(address, debug=True)
    print(json.dumps(result, indent=2))