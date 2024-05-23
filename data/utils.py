import json

import requests
from conf import API_KEY, BASE_URL


def get_order(id):
    url = f"{BASE_URL}/orders/{id}/"
    headers = {
        'Authorization': f'{API_KEY}',
        'Content-Type': 'application/json'
    }
    response = requests.get(url=url, headers=headers)
    if response.ok:
        return json.loads(response.json())
    raise Exception(response.text)


def get_file(url: str, format: str):
    r = requests.get(url)
    with open('./download.' + format, 'wb') as f:
        f.write(r.content)
        return f
