import requests

API_KEY = "הכנס_כאן_את_המפתח_שלך"
url = "https://api.sunoapi.org/v1/songs/generate"

headers = {
    "Authorization": f"Bearer {API_KEY}",
    "Content-Type": "application/json"
}

data = {
    "prompt": "שיר שמח על קיץ וחוף ים",
    "duration": 30
}

response = requests.post(url, headers=headers, json=data)

print("Status code:", response.status_code)
print("Response:", response.text)
