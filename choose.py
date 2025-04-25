import requests

url = "https://softwaredevelopmentsolution.com"
data = {
    "host": "target.com",
    "method": "GET",
    "path": "/api/private"
}
res = requests.post(url, json=data)
print(res.json())
