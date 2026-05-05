#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Comment out all .format( blocks (including multi-line dict parameters) in callmonstr_v5.py.
"""
with open('/root/callmonstr/callmonstr_v5.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
i = 0
while i < len(lines):
    line = lines[i]
    # If line contains .format(, start commenting block
    if '.format(' in line and ('ws.' in line or 'sheet.' in line or 'cell.' in line):
        # Comment this line
        new_lines.append('# ' + line)
        # Continue commenting until we see a line that contains '})' (end of dict)
        # Also stop if we encounter a line that is not part of the dict (e.g., blank line or less indented)
        j = i + 1
        # Determine indentation of the .format( line
        base_indent = len(line) - len(line.lstrip())
        while j < len(lines):
            next_line = lines[j]
            # If next_line is empty or has less indent than base_indent, stop
            if next_line.strip() == '':
                # blank line may be part of block? We'll stop to be safe.
                break
            next_indent = len(next_line) - len(next_line.lstrip())
            # If next_line has indent <= base_indent and is not a continuation (like '}' or '"'), stop
            if next_indent <= base_indent and not next_line.strip().startswith(('"', "'", '{', '}', ':')):
                break
            # If next_line contains '})', it's the end of the dict, comment it and stop
            if '})' in next_line:
                new_lines.append('# ' + next_line)
                i = j + 1
                break
            else:
                new_lines.append('# ' + next_line)
                j += 1
        else:
            # If we didn't break out of inner while, set i appropriately
            i = j
        continue
    else:
        new_lines.append(line)
        i += 1

with open('/root/callmonstr/callmonstr_v5.py', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Commented all .format( blocks.")