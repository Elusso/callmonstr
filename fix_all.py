#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Fix callmonstr_v5.py:
1. Replace lines with '.format(' with 'pass  # format removed', preserving indentation.
2. Remove all lines that are part of the dict (lines with '{', '}', '})' that follow).
"""
with open('/root/callmonstr/callmonstr_v5.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
skip_until_close = False
paren_depth = 0
for line in lines:
    stripped = line.strip()
    # If we encounter a line with '.format('
    if '.format(' in line and ('ws.' in line or 'sheet.' in line or 'cell.' in line):
        # Replace with pass, preserve indentation
        indent = line[:len(line) - len(line.lstrip())]
        new_lines.append(indent + 'pass  # format removed\n')
        # Now skip all following lines that are part of the dict
        skip_until_close = True
        continue
    if skip_until_close:
        # Skip lines that are part of the dict: they have more indentation than the original line?
        # Simpler: skip until we see a line that contains '})'
        if '})' in line:
            skip_until_close = False
        # Also skip lines that start with '"' or contain ':' (dict entries)
        # But we'll just skip until '})'
        continue
    else:
        new_lines.append(line)

with open('/root/callmonstr/callmonstr_v5.py', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Fixed. Check callmonstr_v5.py")