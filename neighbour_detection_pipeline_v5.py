
import requests
from shapely.geometry import shape, Point
from shapely.ops import unary_union
import json 

PLACES_BASE = "https://api.os.uk/search/places/v1"
NGD_BASE = "https://api.os.uk/features/ngd/ofa/v1/collections"

BUFFER_M = 1.0                     # slack distance for polygon-polygon adjacency test
SEARCH_PAD_M = 45.0                # bbox padding around the target property, in metres
ADDRESS_POINT_SLACK_M = 2.0        # buffer applied to a polygon before checking address points
ADDRESS_MATCH_FALLBACK_M = 6.0     # nearest-address-point fallback distance
LAND_CONTAINMENT_FALLBACK_M = 5.0  # nearest-land-parcel fallback distance (target's own parcel)
MIN_SHARED_BOUNDARY_LENGTH_M = 0.3 # minimum length to count as a real shared edge (not a corner)
PINCH_POINT_CHECK_DISTANCE_M = 2.0 # how close a third parcel's midpoint-distance must be to veto
MIN_THIRD_PARTY_AREA_SQM = 30.0    # PLACEHOLDER — tune against real data, see notes above
MAX_PLAUSIBLE_NEIGHBOUR_DISTANCE_M = 30.0  # sanity cap: candidate address point vs target
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


# ---------------------------------------------------------------------------
# Per-property context — replaces the old module-level `srs` global.
# Created fresh at the start of find_neighbours() and threaded through
# every helper explicitly, so nothing can leak stale state between
# properties in a batch run.
# ---------------------------------------------------------------------------
class OSContext:
    def __init__(self, api_key, srs="EPSG:27700"):
        self.api_key = api_key
        self.srs = srs

    def crs(self):
        return CRS_TEMPLATE.format(self.srs.split(":")[-1])


# ---------------------------------------------------------------------------
# STEP 1 — Geocode the target address
# ---------------------------------------------------------------------------
def geocode_address(ctx, address_text):
    url = f"{PLACES_BASE}/find"
    params = {"query": address_text, "maxresults": 1, "key": ctx.api_key}
    r = requests.get(url, params=params)
    r.raise_for_status()
    data = r.json()
    ctx.srs = data.get("header", {}).get("output_srs", ctx.srs)
    results = data.get("results", [])
    print("Current Address Results : ", results)
    if not results:
        return None
    return results[0]["DPA"]


# ---------------------------------------------------------------------------
# STEP 2 — Fetch polygons by TOID / by bbox, for a given NGD collection
# ---------------------------------------------------------------------------
def get_polygon_by_toid(ctx, collection, toid):
    url = f"{NGD_BASE}/{collection}/items"
    params = {
        "key": ctx.api_key,
        "filter": f"toid='{toid}'",
        "filter-lang": "cql-text",
        "crs": ctx.crs(),
    }
    r = requests.get(url, params=params)
    if r.status_code != 200:
        print(f"[{collection}] by-TOID error body:", r.text)
    r.raise_for_status()
    features = r.json().get("features", [])
    return shape(features[0]["geometry"]) if features else None


def get_nearby_polygons(ctx, collection, x, y, pad=SEARCH_PAD_M):
    bbox = f"{x - pad},{y - pad},{x + pad},{y + pad}"
    url = f"{NGD_BASE}/{collection}/items"
    params = {
        "key": ctx.api_key,
        "bbox": bbox,
        "bbox-crs": ctx.crs(),
        "crs": ctx.crs(),
        "limit": 100,
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
# STEP 3 — Address matching (needed BEFORE adjacency, to merge fragments)
# ---------------------------------------------------------------------------
def get_nearby_address_points(ctx, x, y, radius=SEARCH_PAD_M):
    url = f"{PLACES_BASE}/radius"
    params = {"point": f"{x},{y}", "radius": radius, "key": ctx.api_key}
    r = requests.get(url, params=params)
    r.raise_for_status()
    return [res["DPA"] for res in r.json().get("results", []) if "DPA" in res]


def match_address_to_polygon(polygon, address_points,
                              buffer_m=ADDRESS_POINT_SLACK_M,
                              fallback_max_m=ADDRESS_MATCH_FALLBACK_M):
    """Find address point(s) belonging to this polygon. Tries a buffered
    .covers() test first, then falls back to the single nearest address
    point if it's genuinely close (handles edge-of-polygon placements).
    Returns a list of DPA dicts (possibly empty)."""
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


# ---------------------------------------------------------------------------
# STEP 4 — Merge same-owner land fragments into one curtilage polygon.
#
# This is the key fix: OS NGD often represents one property's garden as
# several adjoining polygons (rear strip, side strip, access sliver, etc.)
# instead of a single parcel. Testing adjacency fragment-vs-fragment can
# find a "real shared boundary" between two slivers that both sit several
# strips away from the actual dwellings they belong to — including across
# streets. Merging by matched UPRN first means every adjacency test after
# this point compares whole-curtilage-to-whole-curtilage.
# ---------------------------------------------------------------------------
def build_curtilage_map(nearby_land, address_points):
    """Returns:
        curtilages: dict uprn -> {"polygon": merged Polygon/MultiPolygon,
                                   "address": address string,
                                   "component_toids": [toid, ...]}
        unassigned: list of land features with no address match at all
                    (slivers/paths/etc.) — kept separately since they're
                    still needed for the pinch-point third-party check.
    """
    groups = {}      # uprn -> {"address":..., "polys":[...], "toids":[...]}
    unassigned = []

    for feat in nearby_land:
        toid = feat["properties"]["toid"]
        poly = shape(feat["geometry"])
        matched = match_address_to_polygon(poly, address_points)

        if not matched:
            unassigned.append(feat)
            continue

        # A fragment usually matches one address; if it matches more than
        # one (rare, e.g. a shared driveway), assign it to the first —
        # log this at the call site if you need to audit shared parcels.
        dpa = matched[0]
        uprn = dpa["UPRN"]
        if uprn not in groups:
            groups[uprn] = {"address": dpa["ADDRESS"], "dpa": dpa, "polys": [], "toids": []}
        groups[uprn]["polys"].append(poly)
        groups[uprn]["toids"].append(toid)

    curtilages = {}
    for uprn, g in groups.items():
        merged = unary_union(g["polys"]) if len(g["polys"]) > 1 else g["polys"][0]
        curtilages[uprn] = {
            "polygon": merged,
            "address": g["address"],
            "dpa": g["dpa"],
            "component_toids": g["toids"],
        }

    return curtilages, unassigned


# ---------------------------------------------------------------------------
# STEP 5 — Real geometric adjacency test
# ---------------------------------------------------------------------------
def shares_real_boundary(poly_a, poly_b, min_shared_length=MIN_SHARED_BOUNDARY_LENGTH_M):
    """True only if the two polygons share an actual boundary LINE (not
    just a single touching corner point), of at least min_shared_length."""
    intersection = poly_a.boundary.intersection(poly_b.boundary)

    if intersection.is_empty:
        return None  # no shared boundary at all

    if intersection.geom_type in ("Point", "MultiPoint"):
        return None  # corner-only contact

    if hasattr(intersection, "geoms"):
        total_len = sum(g.length for g in intersection.geoms if hasattr(g, "length"))
    else:
        total_len = intersection.length

    if total_len < min_shared_length:
        return None

    return intersection  # return the actual shared geometry for further checks


def is_clean_two_party_boundary(shared_segment, target_toids, candidate_toids,
                                  all_land_features,
                                  min_distance_from_third=PINCH_POINT_CHECK_DISTANCE_M,
                                  min_third_party_area=MIN_THIRD_PARTY_AREA_SQM):
    """Reject a shared boundary if a THIRD property's parcel sits right at
    its MIDPOINT — a signature of a multi-way pinch point (3+ plots meeting
    at a corner) rather than a genuine two-property boundary. Small parcels
    (driveway/path slivers) are ignored so they don't wrongly veto a real
    neighbour relationship.

    IMPORTANT: target_toids and candidate_toids must be the FULL list of
    component TOIDs making up each side's merged curtilage (not a UPRN —
    land features only carry a toid). The boundary midpoint sits at
    distance 0 from the target's and candidate's own polygons by
    definition (it's their shared edge), so their own fragments MUST be
    excluded here or every boundary self-vetoes. This was a real
    regression in the v5 merge-by-UPRN rewrite: the old TOID-based skip
    was left comparing TOIDs to UPRNs, which is never true, so it silently
    stopped excluding anything.

    NOTE: this checks against the raw (unmerged) land features on purpose —
    the pinch-point signal is about a third party's footprint sitting in
    the way, which is independent of how that third party's own fragments
    got merged."""
    own_toids = set(target_toids) | set(candidate_toids)
    midpoint = shared_segment.interpolate(0.5, normalized=True)
    for feat in all_land_features:
        toid = feat["properties"]["toid"]
        if toid in own_toids:
            continue
        poly = shape(feat["geometry"])
        if poly.area < min_third_party_area:
            continue  # too small to be a real neighbouring property
        if midpoint.distance(poly) < min_distance_from_third:
            return False, toid
    return True, None


# ---------------------------------------------------------------------------
# STEP 6 — Client's filtering rules
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


def sanity_check_distance(target_x, target_y, candidate_dpa, max_dist=MAX_PLAUSIBLE_NEIGHBOUR_DISTANCE_M):
    """Coarse outlier guard: a genuine boundary-sharing neighbour's address
    point should never be far from the target. Catches cases where a chain
    of merged/adjoining fragments produces a geometrically 'clean' boundary
    between two properties that are, in plain terms, nowhere near each
    other (e.g. backing onto a shared rear lane from different streets)."""
    cand_pt = Point(candidate_dpa["X_COORDINATE"], candidate_dpa["Y_COORDINATE"])
    target_pt = Point(target_x, target_y)
    d = target_pt.distance(cand_pt)
    return d <= max_dist, d


# ---------------------------------------------------------------------------
# MAIN PIPELINE
# ---------------------------------------------------------------------------
def find_neighbours(address_text, api_key, debug=False):
    ctx = OSContext(api_key)

    target = geocode_address(ctx, address_text)
    if not target:
        return {"error": "target address not found"}

    target_x, target_y = target["X_COORDINATE"], target["Y_COORDINATE"]
    target_uprn = target["UPRN"]
    target_building_toid = target.get("TOPOGRAPHY_LAYER_TOID")

    rejected = []
    vetoed = []

    # ---- Layer 1: buildings (no pinch-point veto needed here — a shared
    # WALL between exactly two buildings is essentially always genuine) ----
    target_building_poly = None
    if target_building_toid:
        target_building_poly = get_polygon_by_toid(ctx, "bld-fts-buildingpart-1", target_building_toid)

    nearby_buildings = get_nearby_polygons(ctx, "bld-fts-buildingpart-1", target_x, target_y)

    touching_building = []
    if target_building_poly is not None:
        buffered = target_building_poly.buffer(BUFFER_M)
        for feat in nearby_buildings:
            toid = feat["properties"]["toid"]
            if toid == target_building_toid:
                continue
            poly = shape(feat["geometry"])
            if not buffered.intersects(poly):
                continue
            shared = shares_real_boundary(target_building_poly, poly)
            if shared is None:
                continue
            touching_building.append({"toid": toid, "polygon": poly, "source": "building"})

    # ---- Layer 2: land parcels, MERGED BY ADDRESS before adjacency ----
    nearby_land = get_nearby_polygons(ctx, "lnd-fts-land-1", target_x, target_y)
    address_points = get_nearby_address_points(ctx, target_x, target_y)

    curtilages, unassigned_land = build_curtilage_map(nearby_land, address_points)

    # Establish the target's own merged curtilage. Prefer the address-match
    # result (uprn lookup); fall back to raw point-containment if for some
    # reason the target's own address didn't match any land fragment.
    target_curtilage = curtilages.get(target_uprn, {}).get("polygon")
    if target_curtilage is None:
        fallback_toid, target_curtilage = find_containing_polygon(target_x, target_y, nearby_land)
        if target_curtilage is not None:
            # Register a synthetic single-fragment "curtilage" entry so the
            # pinch-point exclusion below still knows which TOID is the
            # target's own, and doesn't self-veto in this fallback path too.
            curtilages[target_uprn] = {
                "polygon": target_curtilage, "address": target["ADDRESS"],
                "dpa": target, "component_toids": [fallback_toid],
            }
            if debug:
                print("Note: target had no address-matched land fragment; used raw containment fallback.")

    touching_land = []
    if target_curtilage is not None:
        buffered_target = target_curtilage.buffer(BUFFER_M)
        for uprn, c in curtilages.items():
            if uprn == target_uprn:
                continue
            poly = c["polygon"]
            if not buffered_target.intersects(poly):
                continue

            shared = shares_real_boundary(target_curtilage, poly)
            if shared is None:
                continue

            target_toids = curtilages.get(target_uprn, {}).get("component_toids", [])
            clean, blocker_toid = is_clean_two_party_boundary(
                shared, target_toids, c["component_toids"], nearby_land
            )
            if not clean:
                vetoed.append({
                    "uprn": uprn, "address": c["address"],
                    "reason": f"pinch-point veto — third-party parcel {blocker_toid} sits at boundary midpoint",
                })
                continue

            ok, dist = sanity_check_distance(target_x, target_y, c["dpa"])
            if not ok:
                vetoed.append({
                    "uprn": uprn, "address": c["address"],
                    "reason": f"address point {dist:.1f}m from target — exceeds plausible neighbour distance "
                              f"({MAX_PLAUSIBLE_NEIGHBOUR_DISTANCE_M}m); likely a cross-street or merged-fragment artefact",
                })
                continue

            touching_land.append({
                "uprn": uprn, "dpa": c["dpa"], "polygon": poly,
                "component_toids": c["component_toids"], "source": "land",
            })

    if debug:
        print("target_building_toid:", target_building_toid)
        print("target_building_poly is None?", target_building_poly is None)
        print("nearby_buildings count:", len(nearby_buildings))
        print("nearby_land count:", len(nearby_land))
        print("curtilages resolved:", len(curtilages), "| unassigned fragments:", len(unassigned_land))
        print("touching_building count:", len(touching_building))
        print("touching_land count:", len(touching_land))
        print("vetoed count:", len(vetoed))

    neighbours = []
    seen_uprns = set()

    # -- from the building layer, still need address matching --
    for t in touching_building:
        matched = match_address_to_polygon(t["polygon"], address_points)
        if not matched:
            rejected.append({"toid": t["toid"], "source": "building",
                              "reason": "no address point found inside polygon"})
            continue
        for dpa in matched:
            uprn = dpa["UPRN"]
            if uprn == target_uprn or uprn in seen_uprns:
                continue
            seen_uprns.add(uprn)

            ok, dist = sanity_check_distance(target_x, target_y, dpa)
            if not ok:
                vetoed.append({"uprn": uprn, "address": dpa.get("ADDRESS"),
                                "reason": f"address point {dist:.1f}m from target — exceeds plausible distance"})
                continue

            valid, reason = is_valid_neighbour(dpa)
            if valid:
                neighbours.append({
                    "uprn": uprn, "address": dpa["ADDRESS"],
                    "classification": dpa.get("CLASSIFICATION_CODE_DESCRIPTION"),
                    "toid": t["toid"], "source": "building",
                })
            else:
                rejected.append({"toid": t["toid"], "uprn": uprn,
                                  "address": dpa.get("ADDRESS"), "reason": reason})

    # -- from the merged land curtilages, address already resolved --
    for t in touching_land:
        uprn = t["uprn"]
        if uprn == target_uprn or uprn in seen_uprns:
            continue
        seen_uprns.add(uprn)
        dpa = t["dpa"]

        valid, reason = is_valid_neighbour(dpa)
        if valid:
            neighbours.append({
                "uprn": uprn, "address": dpa["ADDRESS"],
                "classification": dpa.get("CLASSIFICATION_CODE_DESCRIPTION"),
                "toid": t["component_toids"][0], "component_toids": t["component_toids"],
                "source": "land",
            })
        else:
            rejected.append({"uprn": uprn, "address": dpa.get("ADDRESS"),
                              "toid": t["component_toids"][0], "reason": reason})

    target_is_flat = target.get("CLASSIFICATION_CODE") in EXCLUDED_RESIDENTIAL_SUBTYPES
    aborted = target_is_flat and len(neighbours) > MAX_NEIGHBOURS_BEFORE_ABORT

    return {
        "target": {
            "uprn": target_uprn,
            "address": target["ADDRESS"],
            "building_toid": target_building_toid,
            "land_uprn_resolved": target_uprn in curtilages,
            "classification": target.get("CLASSIFICATION_CODE_DESCRIPTION"),
        },
        "neighbours": neighbours,
        "rejected_candidates": rejected,
        "vetoed_candidates": vetoed,
        "aborted_multi_dwelling": aborted,
        "neighbour_count": len(neighbours),
    }


# ---------------------------------------------------------------------------
# DEBUG HELPER — inspect WHY a boundary is being vetoed, and see the merged
# curtilage fragments for a given property before tuning thresholds.
# ---------------------------------------------------------------------------
def debug_clean_boundary(target_uprn, candidate_uprn, curtilages, nearby_land_feats,
                          min_distance=PINCH_POINT_CHECK_DISTANCE_M):
    target_poly = curtilages.get(target_uprn, {}).get("polygon")
    cand_poly = curtilages.get(candidate_uprn, {}).get("polygon")

    if target_poly is None or cand_poly is None:
        print("Could not find one or both curtilages in the merged map")
        return

    shared = target_poly.boundary.intersection(cand_poly.boundary)
    if shared.is_empty or shared.geom_type in ("Point", "MultiPoint"):
        print("No real shared boundary between these two.")
        return

    own_toids = set(curtilages[target_uprn]["component_toids"]) | set(curtilages[candidate_uprn]["component_toids"])
    midpoint = shared.interpolate(0.5, normalized=True)
    print(f"Shared boundary length: {shared.length:.2f}m")
    print(f"Target curtilage component TOIDs: {curtilages[target_uprn]['component_toids']}")
    print(f"Candidate curtilage component TOIDs: {curtilages[candidate_uprn]['component_toids']}")
    for feat in nearby_land_feats:
        toid = feat["properties"]["toid"]
        if toid in own_toids:
            continue  # these are the two properties' own parcels, not a third party
        poly = shape(feat["geometry"])
        d = midpoint.distance(poly)
        if d < min_distance:
            print(f"  BLOCKED by {toid} at distance {d:.3f}m — area: {poly.area:.1f} sq m")


def debug_property(address_text, api_key):
    """One-shot diagnostic: geocode, fetch land, build curtilage map, and
    print the full touching-candidate table with reasons for every
    accept/reject/veto decision. Use this instead of ad hoc snippets so
    every run is guaranteed to use fresh, correctly-scoped coordinates."""
    ctx = OSContext(api_key)
    target = geocode_address(ctx, address_text)
    if not target:
        print("Target address not found")
        return

    target_x, target_y = target["X_COORDINATE"], target["Y_COORDINATE"]
    nearby_land = get_nearby_polygons(ctx, "lnd-fts-land-1", target_x, target_y)
    address_points = get_nearby_address_points(ctx, target_x, target_y)
    curtilages, unassigned = build_curtilage_map(nearby_land, address_points)

    target_uprn = target["UPRN"]
    print(f"Target: {target['ADDRESS']} (UPRN {target_uprn})")
    if target_uprn not in curtilages:
        print("  WARNING: target UPRN not resolved via address-matched land fragments")
        return
    print(f"  Curtilage area: {curtilages[target_uprn]['polygon'].area:.1f} sqm, "
          f"fragments: {curtilages[target_uprn]['component_toids']}")

    target_poly = curtilages[target_uprn]["polygon"]
    buffered = target_poly.buffer(BUFFER_M)

    for uprn, c in curtilages.items():
        if uprn == target_uprn:
            continue
        if not buffered.intersects(c["polygon"]):
            continue
        shared = shares_real_boundary(target_poly, c["polygon"])
        if shared is None:
            continue
        clean, blocker = is_clean_two_party_boundary(
            shared, curtilages[target_uprn]["component_toids"], c["component_toids"], nearby_land
        )
        ok, dist = sanity_check_distance(target_x, target_y, c["dpa"])
        print(f"  {c['address']} | shared {shared.length:.2f}m | "
              f"pinch_clean={clean}{'' if clean else f' (blocked by {blocker})'} | "
              f"dist={dist:.1f}m plausible={ok}")


if __name__ == "__main__":
    import json
    address = "9 Buxton Close Surrey Epsom KT19 8BD"
    API_KEY = "aJARAMAGjcqsJbfGyA9pXOAuYkwhyxHm"
    result = find_neighbours(address, API_KEY, debug=True)
    print(json.dumps(result['neighbours']))
#
#     # or, for the focused diagnostic used to chase the Burnhams Grove case:
#     debug_property("9 Buxton Close, Epsom, KT19 8BD", API_KEY)
