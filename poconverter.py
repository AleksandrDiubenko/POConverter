import pandas as pd
import json
import re
from google.colab import files
from io import StringIO

# --- Helper functions ---

def parse_po_blocks(content):
    blocks = []
    raw_blocks = re.split(r'\n\s*\n', content.strip(), flags=re.MULTILINE)
    for block in raw_blocks:
        lines = block.splitlines()
        msgstr_index = None
        msgctxt = ""
        msgid = ""
        for idx, line in enumerate(lines):
            if line.startswith("msgctxt"):
                m = re.search(r'"(.*)"', line)
                if m:
                    msgctxt = m.group(1)
            elif line.startswith("msgid"):
                m = re.search(r'"(.*)"', line)
                if m:
                    msgid = m.group(1)
            elif line.startswith("msgstr"):
                msgstr_index = idx
        blocks.append({
            "lines": lines,
            "msgstr_index": msgstr_index,
            "msgctxt": msgctxt,
            "msgid": msgid
        })
    return blocks

def extract_msgstrs(blocks):
    msgstrs = []
    for b in blocks:
        idx = b["msgstr_index"]
        if idx is not None:
            m = re.search(r'"(.*)"', b["lines"][idx])
            msgstrs.append(m.group(1) if m else "")
        else:
            msgstrs.append("")
    return msgstrs

# --- PHASE 1: Convert .po files into Excel ---

def generate_excel_from_pos():
    print("🔹 Phase 1: Upload one or more .po files (same structure).")
    uploaded = files.upload()
    if not uploaded:
        raise ValueError("❌ No files uploaded.")
    
    po_files = {fname: uploaded[fname] for fname in uploaded if fname.lower().endswith('.po')}
    if not po_files:
        raise ValueError("❌ No .po files found in uploaded files.")
    
    contents_data = {}
    tech_rows = []

    file_names = list(po_files.keys())
    print(f"🗂 Found {len(file_names)} .po files.")

    for idx, fname in enumerate(file_names):
        content = po_files[fname].decode("utf-8", errors="replace")
        blocks = parse_po_blocks(content)
        line_count = len(content.splitlines())

        if idx == 0:
            # First file: create ID and SOURCE TEXT columns
            contents_data["ID"] = [b["msgctxt"] for b in blocks]
            contents_data["SOURCE TEXT"] = [b["msgid"] for b in blocks]
            num_blocks = len(blocks)
        
        else:
            # Validate block count consistency
            if len(blocks) != num_blocks:
                raise ValueError(f"❌ File {fname} has {len(blocks)} blocks, expected {num_blocks}.")

        # Add msgstrs to contents_data under this file’s name
        contents_data[fname] = extract_msgstrs(blocks)

        # Store full technical info block-by-block
        for i, blk in enumerate(blocks):
            template = {
                "lines": blk["lines"],
                "msgstr_index": blk["msgstr_index"]
            }
            tech_rows.append({
                "File Name": fname,
                "Block Index": i,
                "Block Template": json.dumps(template, ensure_ascii=False),
                "Line Count": line_count
            })

    # Create Excel file
    df_contents = pd.DataFrame(contents_data)
    df_technical = pd.DataFrame(tech_rows)

    output_excel = "compiled_po_data.xlsx"
    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        df_contents.to_excel(writer, sheet_name="Contents", index=False)
        df_technical.to_excel(writer, sheet_name="Technical", index=False)

    print("✅ Excel created. Downloading...")
    files.download(output_excel)

# --- PHASE 2: Rebuild .po files from Excel ---

def reconstruct_pos_from_excel():
    print("🔹 Phase 2: Upload the Excel file to reconstruct .po files.")
    uploaded = files.upload()
    if not uploaded:
        raise ValueError("❌ No Excel file uploaded.")
    
    excel_fname = next(iter(uploaded))
    df_contents = pd.read_excel(excel_fname, sheet_name="Contents")
    df_technical = pd.read_excel(excel_fname, sheet_name="Technical")

    num_blocks = df_contents.shape[0]
    all_files = df_contents.columns[2:]  # skip ID and SOURCE TEXT
    generated_files = []

    for fname in all_files:
        print(f"\n📄 Reconstructing: {fname}")
        file_blocks = df_technical[df_technical["File Name"] == fname].sort_values(by="Block Index")
        if file_blocks.shape[0] != num_blocks:
            raise ValueError(f"❌ Block mismatch in {fname}: expected {num_blocks}, got {file_blocks.shape[0]}")

        po_blocks = []
        for i, row in file_blocks.iterrows():
            template = json.loads(row["Block Template"])
            msgstr_line_index = template["msgstr_index"]
            lines = template["lines"]
            new_lines = lines.copy()
            new_msgstr = df_contents.iloc[int(row["Block Index"])][fname]
            new_msgstr = "" if pd.isna(new_msgstr) else str(new_msgstr)


            if msgstr_line_index is not None and msgstr_line_index < len(new_lines):
                new_lines[msgstr_line_index] = f'msgstr "{new_msgstr}"'

            po_blocks.append("\n".join(new_lines))

        full_po = "\n\n".join(po_blocks) + "\n"
        with open(fname, "w", encoding="utf-8") as f:
            f.write(full_po)

        actual_lines = full_po.count("\n") + 1
        expected_lines = int(file_blocks["Line Count"].iloc[0])
        print(f"✅ {fname}: {num_blocks} blocks, {actual_lines} lines (expected: {expected_lines})")
        generated_files.append(fname)

    for f in generated_files:
        files.download(f)

# --- Menu ---

print("👋 Welcome! What would you like to do?")
print("1️⃣  Convert .po files ➜ Excel")
print("2️⃣  Reconstruct .po files from Excel")
choice = input("Enter 1 or 2: ").strip()

if choice == "1":
    generate_excel_from_pos()
elif choice == "2":
    reconstruct_pos_from_excel()
else:
    print("❌ Invalid choice.")
