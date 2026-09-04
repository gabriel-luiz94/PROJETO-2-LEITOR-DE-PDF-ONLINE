import json
import urllib.request
with open('static/resumo.js', 'r', encoding='utf-8') as f:
    code = f.read()
data = json.dumps({'code': code}).encode('utf-8')
req = urllib.request.Request('https://esprima.org/demo/validate.js', data=data, headers={'Content-Type': 'application/json'})
try:
    with urllib.request.urlopen(req) as response:
        print(response.read().decode('utf-8'))
except Exception as e:
    print('Failed to use online parser:', e)
