import requests

url = "https://api.green-api.com/waInstance1101751567/checkWhatsapp/f87ed7323cfe45d1b716475bf9210a06de54269a59e344f1af"

payload = "{\r\n    \"phoneNumber\": \r\n}"
headers = {
  'Content-Type': 'application/json'
}

response = requests.request("POST", url, headers=headers, data = payload)

print(response.text)