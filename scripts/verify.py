"""Explicit verification commands; no downloads or automatic mailbox reads.

Run --help for health, local Outlook, model planning and live-chat checks.
"""
import argparse
import asyncio
import json
import os
from pathlib import Path
import sys
import time

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'backend'))


def print_json(value):
    print(json.dumps(value, ensure_ascii=False, indent=2), flush=True)


async def check_outlook():
    from app.outlook import connect_outlook

    async with connect_outlook() as tools:
        definitions = await tools.discover()
        status, _ = await tools.execute('outlook_status', {})
        print_json({'tools': [d['function']['name'] for d in definitions], 'status': status})


async def check_plan(text):
    from app.config import Settings
    from app.llm import Ollama
    from app.outlook import plan_mail

    started = time.monotonic()
    plan = await plan_mail(Ollama(Settings()), text)
    print_json({'plan': plan.model_dump(), 'seconds': round(time.monotonic() - started, 1),
                'mailbox_read': False, 'note': 'Inspect every field; schema validity is not semantic correctness.'})


def check_http(base, query=None):
    import httpx

    token = os.getenv('AGENT_API_TOKEN')
    headers = {'Authorization': 'Bearer ' + token} if token else {}
    with httpx.Client(base_url=base, timeout=270, headers=headers, trust_env=False) as client:
        if query is None:
            response = client.get('/api/health')
            response.raise_for_status()
            health = response.json()
            print_json({key: health.get(key) for key in ('application', 'model', 'ollama', 'speech')})
            print_json({'company_mcp': health.get('mcp', {}).get('available')})
            return
        client.get('/').raise_for_status()
        response = client.post('/api/conversations')
        response.raise_for_status()
        key = response.json()['id']
        response = client.post('/api/chat', json={'conversation_id': key, 'text': query})
        response.raise_for_status()
        events = [json.loads(line) for line in response.text.splitlines() if line]
        for event in events:
            # Never print email bodies/results during a routine verification.
            print_json({key: value for key, value in event.items()
                        if key in {'stage', 'status', 'tool', 'code', 'elapsed_ms', 'operation'}})
        if not events or events[-1].get('stage') != 'final':
            raise RuntimeError('No successful final response; inspect the UI/logs.')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('mode', choices=['health', 'outlook', 'plan', 'chat'])
    parser.add_argument('--text', help='Required for plan/chat. Chat executes the requested real tools.')
    parser.add_argument('--base', default='http://127.0.0.1:8570')
    args = parser.parse_args()
    if args.mode in {'plan', 'chat'} and not args.text:
        parser.error('--text is required')
    if args.mode == 'outlook':
        asyncio.run(check_outlook())
    elif args.mode == 'plan':
        asyncio.run(check_plan(args.text))
    else:
        check_http(args.base, args.text if args.mode == 'chat' else None)


if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        print_json({'success': False, 'code': getattr(exc, 'code', type(exc).__name__),
                    'message': str(exc)})
        sys.exit(1)
