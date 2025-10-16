!pip install xlsxwriter --quiet
import pandas as pd
import re
from google.colab import files
import ipywidgets as widgets
from IPython.display import display, clear_output
import datetime

# --- Helper functions ---

def parse_po_blocks(content):
    blocks = []
    raw_blocks = re.split(r'\n\s*\n', content.strip(), flags=re.MULTILINE)
    for block in raw_blocks:
        lines = block.splitlines()
        msgstr_index = None
        msgctxt = ""
        msgid = ""
        collecting_msgid = False
        collecting_msgctxt = False
        full_msgid = ""
        full_msgctxt = ""
        for idx, line in enumerate(lines):
            if line.startswith("msgctxt"):
                collecting_msgctxt = True
                full_msgctxt = extract_quoted_text(line)
                msgctxt = full_msgctxt
            elif collecting_msgctxt and line.startswith('"'):
                full_msgctxt += extract_quoted_text(line)
                msgctxt = full_msgctxt
            else:
                collecting_msgctxt = False

            if line.startswith("msgid"):
                collecting_msgid = True
                full_msgid = extract_quoted_text(line)
                msgid = full_msgid
            elif collecting_msgid and line.startswith('"'):
                full_msgid += extract_quoted_text(line)
                msgid = full_msgid
            else:
                collecting_msgid = False

            if line.startswith("msgstr"):
                msgstr_index = idx

        blocks.append({
            "lines": lines,
            "msgstr_index": msgstr_index,
            "msgctxt": msgctxt,
            "msgid": msgid
        })
    return blocks

def extract_quoted_text(line):
    m = re.search(r'"(.*)"', line)
    return m.group(1) if m else ""

def extract_msgstrs(blocks, visible_indices):
    msgstrs = []
    for i in visible_indices:
        b = blocks[i]
        idx = b["msgstr_index"]
        if idx is None:
            msgstrs.append("")
            continue

        msgstr_lines = []
        started = False
        for line in b["lines"][idx:]:
            if line.startswith("msgstr"):
                started = True
                m = re.match(r'msgstr\s+"(.*)"', line)
                msgstr_lines.append(m.group(1) if m else "")
            elif started and line.strip().startswith('"'):
                msgstr_lines.append(extract_quoted_text(line))
            else:
                break

        full_msgstr = "\n".join(msgstr_lines)
        msgstrs.append(full_msgstr)
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
    preserve_index = []
    visible_indices = []

    sep = "<|LINE|>"  # safe delimiter

    file_names = list(po_files.keys())
    print(f"🗂 Found {len(file_names)} .po files.")

    num_blocks = None

    for idx, fname in enumerate(file_names):
        content = po_files[fname].decode("utf-8", errors="replace")
        blocks = parse_po_blocks(content)
        line_count = len(content.splitlines())

        if idx == 0:
            id_list = []
            src_texts = []
            contexts = []

            for i, b in enumerate(blocks):
                if b["msgid"].strip() == "" and not any(l.strip().startswith('"') for l in b["lines"] if "msgid" in l):
                    preserve_index.append(i)
                else:
                    visible_indices.append(i)

                    comment_lines = [line.strip() for line in b["lines"] if line.strip().startswith("#")]
                    context_text = "\n".join(comment_lines)

                    contexts.append(context_text)
                    id_list.append(b["msgctxt"])
                    src_texts.append(b["msgid"])

            contents_data["Context"] = contexts
            contents_data["ID"] = id_list
            contents_data["SOURCE TEXT"] = src_texts
            num_blocks = len(blocks)
        else:
            if len(blocks) != num_blocks:
                raise ValueError(f"❌ File {fname} has {len(blocks)} blocks, expected {num_blocks}.")

        all_msgstrs = extract_msgstrs(blocks, list(range(num_blocks)))
        filtered_msgstrs = [all_msgstrs[i] for i in range(num_blocks) if i not in preserve_index]
        contents_data[fname] = filtered_msgstrs

        for i, blk in enumerate(blocks):
            tech_rows.append({
                "File Name": fname,
                "Block Index": i,
                "Block Template": sep.join(blk["lines"]),
                "Line Count": line_count,
                "Visible": i not in preserve_index
            })

    # --- NEW: Check for and handle strings that exceed Excel's cell limit ---
    max_excel_chars = 32767
    problematic_cells = []

    # Check all data that's about to be written to the "Contents" sheet
    for col_name, data_list in contents_data.items():
        for row_index, cell_value in enumerate(data_list):
            if isinstance(cell_value, str) and len(cell_value) > max_excel_chars:
                problematic_cells.append( (col_name, row_index, len(cell_value)) )
                # Truncate the value and add an annotation to prevent the crash
                contents_data[col_name][row_index] = cell_value[:max_excel_chars - 100] + f"... [TRUNCATED - Original length was {len(cell_value)} characters, exceeding Excel's limit]"

    # Print a warning for the user
    if problematic_cells:
        print("\n⚠️  WARNING: Some translation strings exceed Excel's cell limit of 32,767 characters.")
        print("   These cells have been truncated to avoid a fatal error.")
        print("   The problematic cells are (Column, Row Index, Original Length):")
        for col, row, length in problematic_cells:
            print(f"   - {col}, Row ~{row+2}: {length} chars")
        print("   Review these strings in the original .po file. They are likely not single lines but large blocks of text.\n")
    # --- END OF NEW CODE ---

    df_contents = pd.DataFrame(contents_data)
    df_technical = pd.DataFrame(tech_rows)

    output_excel = f"compiled_po_data_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        df_contents.to_excel(writer, sheet_name="Contents", index=False)
        df_technical.to_excel(writer, sheet_name="Technical", index=False)

        workbook  = writer.book
        worksheet = writer.sheets["Contents"]

        for col_num, column in enumerate(df_contents.columns):
            worksheet.set_column(col_num, col_num, 40)

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

    sep = "<|LINE|>"
    all_files = df_contents.columns[3:]
    generated_files = []

    for fname in all_files:
        print(f"\n📄 Reconstructing: {fname}")
        file_blocks = df_technical[df_technical["File Name"] == fname].sort_values(by="Block Index")
        if file_blocks.empty:
            print(f"⚠️ Skipping {fname}: No data found.")
            continue

        po_blocks = []
        translation_index = 0

        for _, row in file_blocks.iterrows():
            lines = str(row["Block Template"]).split(sep)
            msgstr_index = None
            for idx, line in enumerate(lines):
                if line.startswith("msgstr"):
                    msgstr_index = idx
                    break

            visible = row["Visible"]
            new_lines = lines.copy()

            if visible:
                new_msgstr = df_contents.iloc[translation_index][fname]
                translation_index += 1
                if pd.isna(new_msgstr):
                    new_msgstr = ""

                if msgstr_index is not None and 0 <= msgstr_index < len(new_lines):
                    translated_lines = new_msgstr.replace("\r\n", "\n").split("\n")
                    new_msgstr_block = [f'msgstr "{translated_lines[0].replace("\"", "\\\"")}"']
                    for extra_line in translated_lines[1:]:
                        new_msgstr_block.append(f'"{extra_line.replace("\"", "\\\"")}"')

                    end_index = msgstr_index + 1
                    while end_index < len(new_lines) and new_lines[end_index].strip().startswith('"'):
                        end_index += 1
                    new_lines[msgstr_index:end_index] = new_msgstr_block

            po_blocks.append("\n".join(new_lines))

        full_po = "\n\n".join(po_blocks)
        with open(fname, "w", encoding="utf-8") as f:
            f.write(full_po)

        actual_lines = full_po.count("\n") + 1
        expected_lines = int(file_blocks["Line Count"].iloc[0])
        print(f"✅ {fname}: {actual_lines} lines (expected: {expected_lines})")
        generated_files.append(fname)

    for f in generated_files:
        files.download(f)

# --- Menu ---

def start_menu():
    output = widgets.Output()

    def on_button_click(choice):
        button1.disabled = True
        button2.disabled = True
        clear_output(wait=True)
        if choice == "1":
            generate_excel_from_pos()
        elif choice == "2":
            reconstruct_pos_from_excel()

    button1 = widgets.Button(description="📥 .PO ➜ XLSX")
    button2 = widgets.Button(description="📤 XLSX ➜ .PO")

    button1.on_click(lambda b: on_button_click("1"))
    button2.on_click(lambda b: on_button_click("2"))

    display(widgets.VBox([
        widgets.Label("👋 Welcome! Please choose an option:"),
        button1,
        button2
    ]))

start_menu()
