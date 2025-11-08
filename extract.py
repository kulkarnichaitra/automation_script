import os
import re
from pathlib import Path
import pandas as pd


def get_all_py_files(root_dir):
    """Recursively collect all .py files under the given directory."""
    py_files = []
    for root, _, files in os.walk(root_dir):
        for file in files:
            if file.endswith(".py"):
                py_files.append(os.path.join(root, file))
    return py_files


def extract_dict_authors(text):
    mappings = {}
    # find dict keys like "c123456":  or 'c123456':
    key_pattern = re.compile(r'["\'](c\d{6,7})["\']\s*:\s*', re.IGNORECASE)

    # find triple-quoted strings (either """...""" or '''...'''), with optional raw prefix r or R
    triple_q_pattern = re.compile(r'(?:r?"""(.*?)"""|r?\'\'\'(.*?)\'\'\')', re.DOTALL | re.IGNORECASE)

    author_pattern = re.compile(r'@Author\s*[:\-]?\s*(.+)', re.IGNORECASE)

    for m in key_pattern.finditer(text):
        cid = m.group(1)
        # take a reasonable slice after the colon to find the triple-quoted value
        start = m.end()
        tail = text[start:start + 4000]  # search up to 4000 chars after the key (adjust if needed)
        tq = triple_q_pattern.search(tail)
        if not tq:
            continue
        # triple_q_pattern has two capturing groups (one for """..""" and one for '''..''')
        block = tq.group(1) if tq.group(1) is not None else tq.group(2)
        if not block:
            continue
        a = author_pattern.search(block)
        if a:
            author = a.group(1).strip().strip(' :\t\n"\'')
            mappings[cid] = author

    return mappings


def find_author_in_triple_blocks(text, caseid):
    """
    Search all triple-quoted blocks; if a block contains the caseid text and also an @Author,
    return that author. This is a fallback for parametrized caseids.
    """
    triple_q_pattern = re.compile(r'(?:r?"""(.*?)"""|r?\'\'\'(.*?)\'\'\')', re.DOTALL | re.IGNORECASE)
    author_pattern = re.compile(r'@Author\s*[:\-]?\s*(.+)', re.IGNORECASE)
    for m in triple_q_pattern.finditer(text):
        block = m.group(1) if m.group(1) is not None else m.group(2)
        if not block:
            continue
        if caseid in block:
            a = author_pattern.search(block)
            if a:
                return a.group(1).strip().strip(' :\t\n"\'')
    return ""


def extract_author_from_nearby_block(lines, func_idx):
    """
    For inline tests: scan forward from the function def (func_idx) up to a window
    to find a triple-quoted block or standalone @Author and return the author found.
    """
    max_lines = 60  # adjust if needed; limits how far we look after the def
    joined = "\n".join(lines[func_idx:func_idx + max_lines])
    triple_q_pattern = re.compile(r'(?:r?"""(.*?)"""|r?\'\'\'(.*?)\'\'\')', re.DOTALL | re.IGNORECASE)
    author_pattern = re.compile(r'@Author\s*[:\-]?\s*(.+)', re.IGNORECASE)

    m = triple_q_pattern.search(joined)
    if m:
        block = m.group(1) if m.group(1) is not None else m.group(2)
        if block:
            a = author_pattern.search(block)
            if a:
                return a.group(1).strip().strip(' :\t\n"\'')
    # fallback: scan lines for standalone @Author
    for i in range(func_idx, min(len(lines), func_idx + max_lines)):
        a_line = re.search(r'@Author\s*[:\-]?\s*(.+)', lines[i], re.IGNORECASE)
        if a_line:
            return a_line.group(1).strip().strip(' :\t\n"\'')
    return ""


def create_caseid_file(py_files, out_xlsx="CaseId_Results.xlsx"):
    results = []

    # Patterns
    func_pattern = re.compile(r'^\s*def\s+(test_[a-zA-Z0-9_]+)\s*\(', re.IGNORECASE)
    inline_caseid_pattern = re.compile(r'(c\d{6,7})', re.IGNORECASE)
    param_caseid_pattern = re.compile(r'pytest\.param\(\s*["\'](c\d{6,7})["\']', re.IGNORECASE)
    param_decorator_start = re.compile(r'^\s*@pytest\.mark\.parametrize', re.IGNORECASE)

    for file_path in py_files:
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                text = f.read()
        except Exception as e:
            print(f"Error reading file ({file_path}): {e}")
            continue

        # Build mapping from dict keys only (no @testcaseId usage)
        file_caseid_author_map = extract_dict_authors(text)
        lines = text.splitlines()
        decorator_block = []

        for idx, raw_line in enumerate(lines):
            line_strip = raw_line.strip()

            # Start of parametrize decorator block
            if param_decorator_start.match(line_strip):
                decorator_block = [line_strip]
                continue

            # Continuation of decorator block until function def
            if decorator_block and not line_strip.startswith('def') and line_strip:
                decorator_block.append(line_strip)
                continue

            # Function definition encountered
            func_match = func_pattern.match(line_strip)
            if func_match:
                func_name = func_match.group(1)

                # Inline test: caseid from function name; author from nearest block
                inline_caseid_match = inline_caseid_pattern.search(func_name)
                if inline_caseid_match:
                    caseid = inline_caseid_match.group(1)
                    author = extract_author_from_nearby_block(lines, idx)
                    results.append({
                        "CaseId": caseid,
                        "Automation Script": func_name,
                        "Author": author
                    })

                # Parametrized test: caseids from decorator; author from dict keys mapping OR fallback triple block search
                elif decorator_block:
                    decorator_text = " ".join(decorator_block)
                    param_caseids = param_caseid_pattern.findall(decorator_text)
                    for caseid in param_caseids:
                        # 1) try dict-key mapping
                        author = file_caseid_author_map.get(caseid, "")
                        # 2) fallback: search triple-quoted blocks that contain the caseid text
                        if not author:
                            author = find_author_in_triple_blocks(text, caseid)
                        results.append({
                            "CaseId": caseid,
                            "Automation Script": f"{func_name}[{caseid}]",
                            "Author": author
                        })

                # reset decorator block after function handled
                decorator_block = []

        # end per-file

    # Write results to Excel (keep same column order)
    if results:
        df = pd.DataFrame(results, columns=["CaseId", "Automation Script", "Author"])
        df.to_excel(out_xlsx, index=False)
        print(f"Excel '{out_xlsx}' created with {len(results)} entries.")
    else:
        print("No matching CaseIds found.")


def main():
    directory = Path(r"D:\Pytest\tests")
    if not os.path.isdir(directory):
        print("Folder path is invalid.")
        return

    py_files = get_all_py_files(directory)
    if not py_files:
        print("No python files found in directory.")
        return

    create_caseid_file(py_files)


if __name__ == "__main__":
    main()
