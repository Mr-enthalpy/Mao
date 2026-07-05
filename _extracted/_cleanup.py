#!/usr/bin/env python3
"""Clean up end_notes: remove non-【需核】 items across all month scripts."""
import os, re

BASE = r'E:\1952年大传\1952年大传\_extracted'

for m in ['1952.3','1952.4','1952.5','1952.6','1952.7','1952.8','1952.9','1952.10','1952.11','1952.12']:
    path = os.path.join(BASE, m, 'batch_annotate.py')
    if not os.path.exists(path):
        continue
    
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    lines = content.split('\n')
    new_lines = []
    
    for line in lines:
        # Skip lines with non-需核 labels inside end_notes
        stripped = line.strip()
        if any(tag in stripped for tag in ['【可能过注】', '【需统一】', '【建议改写】']):
            continue
        new_lines.append(line)
    
    with open(path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(new_lines))
    
    print('Done:', m)

print('All cleaned.')
