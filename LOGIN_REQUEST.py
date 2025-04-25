import requests
from concurrent.futures import ThreadPoolExecutor, as_completed

# Set login URL (update this if actual endpoint differs)
LOGIN_URL = "https://softwaredevelopmentsolution.com/login"  # Adjust if needed

# Payload based on assumed form fields
payload = {
    "email": "ola123@yopmail.com",
    "password": "1"
}

# Optional headers, modify if your endpoint requires specific headers
headers = {
    "Content-Type": "application/x-www-form-urlencoded"
}

# Total number of requests to simulate
TOTAL_REQUESTS = 1000

def send_login_request(request_id):
    try:
        response = requests.post(LOGIN_URL, data=payload, headers=headers, timeout=10)
        return (request_id, response.status_code, response.url)
    except Exception as e:
        return (request_id, "ERROR", str(e))

def main():
    with ThreadPoolExecutor(max_workers=100) as executor:
        futures = [executor.submit(send_login_request, i) for i in range(TOTAL_REQUESTS)]

        for future in as_completed(futures):
            request_id, status, result = future.result()
            print(f"[{request_id}] Status: {status} | Result: {result}")

if __name__ == "__main__":
    main()
