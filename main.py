import requests
import json

url = 'https://lselectricdwp.webhook.office.com/webhookb2/67573c49-c614-4240-908a-bc778b3855a9@247258cc-5eb2-4fd4-9bb2-f272103f0c34/IncomingWebhook/a2f13b2f58de4f7eac433fb7f5a73e63/667762b1-4f18-4ffe-bf9c-3b0df246edc0'

payload = {
    "text": "Sample alert text"
}
headers = {
    'Content-Type': 'application/json'
}

response = requests.post(url, headers=headers, data=json.dumps(payload))

print(response.text.encode('utf8'))