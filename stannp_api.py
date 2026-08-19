import requests
import base64

url = "https://api-eu1.stannp.com/v1/letters/post"

api_key = "9d0f3cf6a86ca55a00b13e9b"

auth = base64.b64encode(f"{api_key}:".encode()).decode()

payload = {
    "test": "1",
    "country": "GB",
    "size": "A4",
    "duplex": "true"
}

files = {
    "pdf": (
        "letter_one_draft.pdf",
        open(r"c:\Users\hp\Downloads\letter_one_draft.pdf", "rb"),
        "application/pdf"
    ),   
}

headers = {
    "Authorization": f"Basic {auth}"
}

response = requests.post(
    url,
    headers=headers,
    data=payload,
    files=files
)

print(response.status_code)
print(response.text)
