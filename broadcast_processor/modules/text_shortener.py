import requests
import json
import yaml

def load_config():
    config_path = "config.yaml"
    try:
        with open(config_path, 'r', encoding='utf-8') as file:
            config = yaml.safe_load(file)
            return config.get('gpt', {})
    except Exception:
        return {}

def shorten_text(text):
    config = load_config()
    if not config:
        return "Ошибка: не удалось загрузить конфигурацию"
    api_key = config.get('api_key')
    proxy_url = config.get('proxy_url')
    max_chars = config.get('max_chars', 200)
    timeout = config.get('timeout', 30)
    if not api_key or not proxy_url:
        return "Ошибка: отсутствуют обязательные параметры в конфигурации"
    system_prompt = f"Сократи данный текст до максимум {max_chars} символов, сохраняя основной смысл. Отвечай только сокращенным текстом без дополнительных комментариев."
    headers = {"Content-Type": "application/json"}
    payload = {
        "api_key": api_key,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": text}
        ]
    }
    try:
        response = requests.post(proxy_url, json=payload, headers=headers, timeout=timeout)
        if response.status_code == 200:
            api_response = response.json()
            return api_response.get("response", "Пустой ответ от API")
        else:
            return f"Ошибка API: {response.status_code} - {response.text}"
    except Exception as e:
        return f"Ошибка: {str(e)}"
