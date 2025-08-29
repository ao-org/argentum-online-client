import os
import re

# File types to process
VB_EXTENSIONS = ('.bas', '.cls', '.frm')

# Regex to find Sub/Function declarations
SUB_FUNC_PATTERN = re.compile(
    r'^(?P<indent>\s*)(Public|Private)?\s*(Sub|Function)\s+(?P<name>\w+)\b.*$', re.IGNORECASE | re.MULTILINE
)

def add_error_handling_to_code(code, modulename):
    lines = code.splitlines(keepends=True)
    output = []
    i = 0
    while i < len(lines):
        line = lines[i]
        match = SUB_FUNC_PATTERN.match(line)
        if match:
            indent = match.group('indent')
            subfunc_type = match.group(3)
            name = match.group('name')
            # Look ahead for On Error already present
            on_error_line = i + 1 < len(lines) and 'On Error Goto' in lines[i+1]
            if on_error_line:
                output.append(line)
                i += 1
                continue  # Already has error handling
            # Insert On Error Goto after the declaration
            output.append(line)
            output.append(f"{indent}    On Error Goto {name}_Err\r\n")
            # Find where End Sub/Function is
            j = i+1
            end_found = False
            while j < len(lines):
                if re.match(rf'^{indent}End {subfunc_type}\b', lines[j].strip(), re.IGNORECASE):
                    # Insert Exit Sub/Function and error handler before End
                    output.append(f"{indent}    Exit {subfunc_type}\r\n")
                    output.append(f"{indent}{name}_Err:\r\n")
                    output.append(
                        f'{indent}    Call TraceError(Err.Number, Err.Description, "{modulename}.{name}", Erl)\r\n'
                    )
                    end_found = True
                    break
                output.append(lines[j])
                j += 1
            if end_found:
                output.append(lines[j])  # End Sub/Function line
                i = j  # Skip to after End
            else:
                # Malformed sub/function, just append rest
                i = j - 1
        else:
            output.append(line)
        i += 1
    return ''.join(output)

def process_file(filepath):
    # Ignore .bak files (for safety, although we won't create them)
    if filepath.lower().endswith('.bak'):
        return
    modulename = os.path.splitext(os.path.basename(filepath))[0]
    with open(filepath, 'r', encoding='cp1252', errors='ignore') as f:
        code = f.read()
    new_code = add_error_handling_to_code(code, modulename)
    if new_code != code:
        # Overwrite original file only, no backup
        with open(filepath, 'w', encoding='cp1252') as f:
            f.write(new_code)
        print(f'Processed: {filepath}')
    else:
        print(f'No changes: {filepath}')

def main():
    root = os.getcwd()
    for dirpath, _, filenames in os.walk(root):
        for filename in filenames:
            # Ignore .bak files entirely
            if filename.lower().endswith('.bak'):
                continue
            if filename.lower().endswith(VB_EXTENSIONS):
                filepath = os.path.join(dirpath, filename)
                process_file(filepath)

if __name__ == '__main__':
    main()