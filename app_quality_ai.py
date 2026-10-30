from copy import deepcopy

def create_wi_fixed(template_file, doc_no, product, temp_no, rev, proc, part_req, tools_req, time_df, steps_text, approval_di_bawah=True):
    # Load template ASLI - biar style, logo, font tetep plek
    template_doc = Document(template_file)
    
    # Simpan 2 tabel asli (Approval & History) persis apa adanya
    saved_tables = []
    for tbl in template_doc.tables:
        saved_tables.append(deepcopy(tbl._tbl))
    
    # Hapus tabel asli dari atas
    for tbl in list(template_doc.tables):
        tbl._tbl.getparent().remove(tbl._tbl)
    
    # Hapus paragraf kosong di atas biar gak banyak spasi
    for p in template_doc.paragraphs:
        if not p.text.strip():
            p.text = ""
    
    doc = template_doc # pakai doc yang sama biar style & landscape tetep
    
    # === FUNGSI TAMBAH WI (di atas atau di bawah tergantung opsi) ===
    def add_wi_content(doc):
        # Judul
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f"WORK INSTRUCTION\n{proc}\nProduct: {product}\nDocument No: {doc_no}")
        run.bold = True
        run.font.size = Pt(12)
        
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f"Doc: {doc_no} | Temp: {temp_no} | Rev: {rev} | Date: {datetime.now().strftime('%d/%m/%Y')}")
        run.font.size = Pt(8)
        
        p = doc.add_paragraph()
        run = p.add_run("Langkah Kerja")
        run.bold = True
        
        t = doc.add_table(rows=1, cols=2)
        t.style = 'Table Grid'
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.cell(0,0).text = "Langkah Kerja"
        t.cell(0,1).text = "Time (min)"
        t.cell(0,0).paragraphs[0].runs[0].bold = True
        t.cell(0,1).paragraphs[0].runs[0].bold = True
        
        total = 0
        for _, r in time_df.iterrows():
            rw = t.add_row().cells
            rw[0].text = str(r['step'])
            rw[1].text = str(r['time'])
            try: total += int(r['time'])
            except: pass
        rw = t.add_row().cells
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
        if not steps:
            steps = [("Steps", [x for x in steps_text.split('\n') if x.strip()])]
        
        for title, bullets in steps:
            p = doc.add_paragraph()
            r = p.add_run(title)
            r.bold = True
            for b in bullets:
                if b:
                    doc.add_paragraph(f"●  {b}")

    def add_tables_from_template(doc):
        # Tempel balik tabel asli plek 100%
        for tbl_element in saved_tables:
            doc._body._element.append(deepcopy(tbl_element))
            doc.add_paragraph("") # spasi antar tabel
        
        # Update isi History (tabel ke-2) setelah ditempel
        try:
            # Tabel history adalah tabel terakhir
            last_table = doc.tables[-1]
            last_table.cell(2,1).text = datetime.now().strftime("%m/%d/%Y")
            last_table.cell(2,2).text = rev
            last_table.cell(2,3).text = f"{proc} - {product}"
        except:
            pass

    # === LOGIKA POSISI - INI YANG KAMU MAU ===
    if approval_di_bawah:
        # WI DULU, BARU APPROVAL DI BAWAH
        add_wi_content(doc)
        doc.add_page_break()
        add_tables_from_template(doc)
    else:
        # APPROVAL DI ATAS (default template)
        add_tables_from_template(doc)
        doc.add_page_break()
        add_wi_content(doc)
    
    return doc
