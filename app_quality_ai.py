def create_wi_landscape_bottom(template_file, doc_no, product, temp_no, rev, proc, part_req, tools_req, time_df, steps_text):
    doc = Document(template_file)  # Langsung pakai template asli biar header + logo kepakai
    
    # 1. Simpan 2 tabel asli
    saved_tables = list(doc.tables)
    
    # 2. Hapus HANYA tabelnya dari atas - JANGAN hapus paragraf (biar logo tidak hilang)
    for tbl in saved_tables:
        tbl._tbl.getparent().remove(tbl._tbl)
    
    # 3. TAMBAH WI DI ATAS (setelah header logo)
    # (kode WI sama seperti sebelumnya - Langkah Kerja, Part, Tools, Steps)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"WORK INSTRUCTION\n{proc}\nProduct: {product}\nDocument No: {doc_no}")
    run.bold = True
    run.font.size = Pt(14)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"Doc: {doc_no} | Temp: {temp_no} | Rev: {rev} | Date: {datetime.now().strftime('%d/%m/%Y')}")
    run.font.size = Pt(8)
    run.bold = True

    p = doc.add_paragraph()
    run = p.add_run("Langkah Kerja")
    run.bold = True
    run.font.size = Pt(11)

    t2 = doc.add_table(rows=1, cols=2)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    t2.style = 'Table Grid'
    t2.cell(0,0).text = "Langkah Kerja"
    t2.cell(0,1).text = "Time (min)"
    t2.cell(0,0).paragraphs[0].runs[0].bold = True
    t2.cell(0,1).paragraphs[0].runs[0].bold = True

    total = 0
    for _, row in time_df.iterrows():
        rw = t2.add_row().cells
        rw[0].text = str(row['step'])
        rw[1].text = str(row['time'])
        try: total += int(row['time'])
        except: pass
    rw = t2.add_row().cells
    rw[0].text = "Total"
    rw[1].text = str(total)
    rw[0].paragraphs[0].runs[0].bold = True
    rw[1].paragraphs[0].runs[0].bold = True

    doc.add_paragraph("")
    p = doc.add_paragraph()
    p.add_run("Part required :").bold = True
    for it in part_req.split('\n'):
        if it.strip(): doc.add_paragraph(f"●  {it.strip()}")
    
    p = doc.add_paragraph()
    p.add_run("Tools required :").bold = True
    for it in tools_req.split('\n'):
        if it.strip(): doc.add_paragraph(f"●  {it.strip()}")

    # Steps
    pattern = r"(Step\s*\d+\s*:.*)"
    parts = re.split(pattern, steps_text, flags=re.I)
    steps = []
    for i in range(1, len(parts), 2):
        title = parts[i].strip()
        body = parts[i+1] if i+1 < len(parts) else ""
        bullets = [x.strip() for x in body.split('\n') if x.strip()]
        steps.append((title, bullets))

    for title, bullets in steps:
        p = doc.add_paragraph()
        p.add_run(title).bold = True
        for b in bullets:
            if b: doc.add_paragraph(f"●  {b}")

    # 4. PINDAHKAN 2 TABEL ASLI KE BAWAH
    doc.add_page_break()
    for tbl in saved_tables:
        doc._body._element.append(tbl._tbl)
        doc.add_paragraph("")

    # Update history
    try:
        hist_table = doc.tables[-1]
        hist_table.cell(2,1).text = datetime.now().strftime("%m/%d/%Y")
        hist_table.cell(2,2).text = rev
        hist_table.cell(2,3).text = f"{proc} - {product}"
    except: pass

    return doc
