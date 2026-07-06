import urllib.request
import urllib.error
import json

base_url = "http://192.168.39.143:8001/v1"

print("Step 1: Testing /v1/models endpoint...")
try:
    req = urllib.request.Request(f"{base_url}/models", method="GET")
    with urllib.request.urlopen(req, timeout=10) as response:
        data = response.read().decode('utf-8')
        json_data = json.loads(data)
        print("Success! Models response:")
        print(json.dumps(json_data, indent=2, ensure_ascii=False))
except urllib.error.URLError as e:
    print(f"Error connecting to models endpoint: {e}")
except Exception as e:
    print(f"General error: {e}")

print("\nStep 2: Testing chat completions endpoint with model 'gemma-4:31B'...")
payload = {
    "model": "gemma-4:31B",
    "messages": [
        {"role": "user", "content": "你好，請用繁體中文自我介紹，並告訴我你擅長什麼事？"}
    ],
    "temperature": 0.2,
    "max_tokens": 500
}

try:
    req_data = json.dumps(payload).encode('utf-8')
    req = urllib.request.Request(
        f"{base_url}/chat/completions",
        data=req_data,
        headers={"Content-Type": "application/json"},
        method="POST"
    )
    with urllib.request.urlopen(req, timeout=30) as response:
        data = response.read().decode('utf-8')
        result = json.loads(data)
        print("Success! Chat completion response:")
        print(json.dumps(result, indent=2, ensure_ascii=False))
        content = result['choices'][0]['message']['content']
        print("\n--- Model Response ---")
        print(content)
        print("----------------------")
except urllib.error.URLError as e:
    print(f"Error connecting to chat completions endpoint: {e}")
except Exception as e:
    print(f"General error: {e}")

