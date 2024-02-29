import requests

url = 'http://172.19.1.21/medpt/cyp2001.php'


data = {
    'g_yyymmdd_s': '113/02/29',
    'from': 'cy',
}

response = requests.post(url, data=data)

with open('d:\\mis\\response.html', 'wb') as f:
    f.write(response.content)


