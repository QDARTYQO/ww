import requests

API_KEY = "cbda1568442d9391191970c30a84142a"
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
