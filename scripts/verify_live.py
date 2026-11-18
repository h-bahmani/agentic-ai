"""Run after the user's model download completes. Never pulls or installs anything."""
import json
import os
import sys
import httpx

def main() -> int:
    base = os.getenv('AGENT_BASE_URL', 'http://127.0.0.1:8570')
    token = os.getenv('AGENT_API_TOKEN')
    with httpx.Client(timeout=270, trust_env=False,
                      headers={'Authorization': 'Bearer ' + token} if token else {}) as client:
        health = client.get(base + '/api/health')
        health.raise_for_status()
        state = health.json()
        print(json.dumps(state, ensure_ascii=False, indent=2))
        if not state['ollama']['available'] or not state['mcp']['available']:
            print('NOT VERIFIED: a required backend/model is unavailable; nothing downloaded.')
            return 2
        client.get(base + '/').raise_for_status()
        conversation = client.post(base + '/api/conversations')
        conversation.raise_for_status()
        response = client.post(base + '/api/chat', json={
            'conversation_id': conversation.json()['id'],
            'text': 'تعداد کل سفارش‌های شرکت چقدر است؟ نتیجه واقعی را بررسی کن و فارسی جواب بده.'})
        response.raise_for_status()
        events = [json.loads(line) for line in response.text.splitlines() if line]
        for event in events: print(json.dumps(event, ensure_ascii=False))
        real = [e for e in events if e.get('tool') == 'sales_order_count' and e.get('status') == 'success']
        final = [e for e in events if e.get('stage') == 'final' and e.get('answer')]
        if not real or not final:
            print('NOT VERIFIED: the model did not complete the real order-count tool loop.')
            return 3
        print('REAL TOOL LOOP COMPLETED. Review the displayed Persian answer against the actual evidence.')
        return 0

if __name__ == '__main__': sys.exit(main())
