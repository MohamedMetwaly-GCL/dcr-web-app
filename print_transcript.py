import json
import sys
sys.stdout.reconfigure(encoding='utf-8')
log_path = r'C:\Users\Mohamed Metwaly\.gemini\antigravity\brain\18cf51b3-c15f-40e9-9b89-9edb67d01440\.system_generated\logs\transcript.jsonl'
for line in open(log_path, encoding='utf-8'):
    if '"type":"USER_INPUT"' in line:
        data = json.loads(line)
        content = data.get('content', '')
        if 'UI' in content or 'تقرير' in content or 'المشكلة' in content or 'ملاحظات' in content or 'المطلوب' in content:
            print('--- USER INPUT ---')
            print(content)
