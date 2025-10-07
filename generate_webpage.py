import openpyxl
import json
import os
import re

# --- Configuration ---
EXCEL_FILE = 'viva_links_fixed.xlsx'
IMAGE_FOLDER_NAME = 'downloaded_images' 
OUTPUT_SCRIPT_FILE = 'script.js'
OUTPUT_HTML_FILE = 'index.html'
OUTPUT_CSS_FILE = 'style.css'
SHEETS_TO_PROCESS = {
    'anat': 'Anat',
    'path': 'Path',
    'physio': 'Physio',
    'pharm': 'Pharm',
    'cbb': 'CBB'
}
SHEET_COLORS = {
    'anat': '#3498db',
    'path': '#e74c3c',
    'physio': '#2ecc71',
    'pharm': '#9b59b6',
    'cbb': '#f39c12'
}

def generate_web_data():
    """Reads the Excel file and extracts all data, including standard and formula-based hyperlinks."""
    print(f"Reading data from '{EXCEL_FILE}'...")
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ Error: Excel file not found at '{EXCEL_FILE}'")
        return None

    workbook = openpyxl.load_workbook(EXCEL_FILE, data_only=False)
    workbook_data_only = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    
    all_subject_data = {}

    for key, sheet_name in SHEETS_TO_PROCESS.items():
        if sheet_name not in workbook.sheetnames:
            print(f"⚠️ Warning: Sheet '{sheet_name}' not found. Skipping.")
            continue
        
        worksheet = workbook[sheet_name]
        worksheet_data_only = workbook_data_only[sheet_name]
        
        headers = [cell.value for cell in worksheet[2]]
        sheet_sections = []
        
        for i, header_name in enumerate(headers):
            if header_name and isinstance(header_name, str):
                sheet_sections.append({ "header": header_name.strip(), "col_index": i, "items": [] })
        
        for row_idx, row in enumerate(worksheet.iter_rows(min_row=3), start=3):
            for section in sheet_sections:
                desc_col_index = section['col_index'] + 1
                status_col_index = section['col_index'] + 2
                
                if len(row) > status_col_index:
                    desc_cell = row[desc_col_index] 
                    status_cell = row[status_col_index]

                    if desc_cell.value:
                        item_data = { "desc": "", "status": str(status_cell.value) if status_cell.value else "Pending", "hyperlink": "" }
                        
                        if desc_cell.hyperlink:
                            item_data["desc"] = str(desc_cell.value)
                            item_data["hyperlink"] = desc_cell.hyperlink.target.replace("\\", "/")
                        elif isinstance(desc_cell.value, str) and desc_cell.value.strip().upper().startswith('=HYPERLINK'):
                            match = re.search(r'=HYPERLINK\("([^"]+)"(?:,\s*"([^"]+)")?\)', desc_cell.value, re.IGNORECASE)
                            if match:
                                link_target = match.group(1)
                                friendly_name_cell = worksheet_data_only.cell(row=row_idx, column=desc_col_index + 1)
                                item_data["desc"] = str(friendly_name_cell.value)
                                item_data["hyperlink"] = link_target.replace("\\", "/")
                            else:
                                friendly_name_cell = worksheet_data_only.cell(row=row_idx, column=desc_col_index + 1)
                                item_data["desc"] = str(friendly_name_cell.value)
                        else:
                            item_data["desc"] = str(desc_cell.value)
                        
                        if item_data["hyperlink"]:
                            item_data["hyperlink"] = item_data["hyperlink"].replace("downloaded images/", f"{IMAGE_FOLDER_NAME}/")
                        
                        section["items"].append(item_data)
        
        all_subject_data[key] = {
            "title": sheet_name,
            "color": SHEET_COLORS.get(key, '#7f8c8d'),
            "data": [s for s in sheet_sections if s["items"]]
        }
    print("✅ Data extraction complete.")
    return all_subject_data

def write_html_file():
    html_content = """<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Viva Study Tracker</title><link rel="stylesheet" href="style.css"></head>
<body>
    <header><h1>Viva Study Tracker</h1><nav><button class="tab-btn active" data-tab="dashboard">Dashboard</button><button class="tab-btn" data-tab="anat">Anat</button><button class="tab-btn" data-tab="path">Path</button><button class="tab-btn" data-tab="physio">Physio</button><button class="tab-btn" data-tab="pharm">Pharm</button><button class="tab-btn" data-tab="cbb">CBB</button></nav></header>
    <main>
        <div id="dashboard" class="tab-content active">
            
            <div class="disclaimer-box">
                <p><strong>Acknowledgement:</strong> All content is sourced from the ACEM training site and EDvivas. This page is for personal data reorganization only.</p>
                <a href="https://notebooklm.google.com/notebook/e1cdd0b3-3f57-452b-84e3-bc3fd476d80c?authuser=3" target="_blank" class="notebook-link">Open Study Notes in NotebookLM</a>
            </div>

            <h2>Overall Progress</h2>
            <div class="progress-container"><div class="progress-bar" id="total-progress-bar"></div><span class="progress-label" id="total-progress-label"></span></div>
            <div class="subject-progress-grid"><div id="anat-progress-card" class="progress-card"></div><div id="path-progress-card" class="progress-card"></div><div id="physio-progress-card" class="progress-card"></div><div id="pharm-progress-card" class="progress-card"></div><div id="cbb-progress-card" class="progress-card"></div></div>
            <div class="downloads-section">
                <h2>Downloads</h2>
                <a href="combined1.pdf" class="download-link" download>Download PDF Summary</a>
                <a href="viva_links_fixed.xlsx" class="download-link" download>Download Excel Source File</a>
            </div>
        </div>
        <div id="anat" class="tab-content"></div><div id="path" class="tab-content"></div><div id="physio" class="tab-content"></div><div id="pharm" class="tab-content"></div><div id="cbb" class="tab-content"></div>
    </main>
    <script src="script.js"></script>
</body>
</html>"""
    with open(OUTPUT_HTML_FILE, 'w', encoding='utf-8') as f: f.write(html_content)
    print(f"✅ Created '{OUTPUT_HTML_FILE}'")

def write_css_file():
    css_content = """body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;background-color:#f4f7f9;color:#333;margin:0;padding:20px}header{text-align:center;margin-bottom:20px}h1{color:#2c3e50}h2{color:#34495e;border-bottom:2px solid #e0e0e0;padding-bottom:10px;margin-top:40px}nav{display:flex;justify-content:center;background-color:#fff;border-radius:8px;padding:5px;box-shadow:0 2px 4px rgba(0,0,0,.1);margin-bottom:20px}.tab-btn{padding:10px 20px;border:none;background-color:transparent;cursor:pointer;font-size:16px;border-radius:6px;transition:background-color .3s,color .3s}.tab-btn.active{background-color:#3498db;color:#fff}.tab-btn:hover:not(.active){background-color:#ecf0f1}.tab-content{display:none;padding:20px;background-color:#fff;border-radius:8px;box-shadow:0 2px 4px rgba(0,0,0,.1)}.tab-content.active{display:block}.progress-container{width:100%;background-color:#e0e0e0;border-radius:25px;margin:15px 0;position:relative;height:30px}.progress-bar{height:100%;background-color:#2ecc71;border-radius:25px;text-align:center;line-height:30px;color:#fff;width:0;transition:width .5s ease-in-out}.progress-label{position:absolute;width:100%;text-align:center;top:0;left:0;line-height:30px;font-weight:700;color:#333}.subject-progress-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(250px,1fr));gap:20px;margin-top:20px}.progress-card{padding:15px;background:#fff;border-left:5px solid #3498db}.progress-card h3{margin-top:0}.subject-table-container{display:grid;grid-template-columns:repeat(auto-fit,minmax(350px,1fr));gap:20px}table{width:100%;border-collapse:collapse;font-size:14px}th,td{padding:10px;border:1px solid #ddd;text-align:left}th{background-color:#ecf0f1}td a{text-decoration:none;color:#2980b9;font-weight:500}td a:hover{text-decoration:underline}.status-done{background-color:#d4edda;color:#155724;font-weight:700}.status-pending{background-color:#fff3cd;color:#856404}.downloads-section{margin-top:40px;padding-top:20px;border-top:2px solid #e0e0e0;text-align:center}.download-link{display:inline-block;margin:5px 10px;padding:10px 18px;background-color:#7f8c8d;color:#fff;text-decoration:none;border-radius:5px;font-weight:700;transition:background-color .3s}.download-link:hover{background-color:#6c7a7d}
/* --- NEW CSS FOR DISCLAIMER AND NOTEBOOK LINK ADDED HERE --- */
.disclaimer-box{background-color:#ecf0f1;border-left:5px solid #7f8c8d;padding:15px;margin-bottom:30px;border-radius:5px;text-align:center}.disclaimer-box p{margin:0 0 15px;color:#555}.notebook-link{display:inline-block;padding:10px 18px;background-color:#4285F4;color:#fff;text-decoration:none;border-radius:5px;font-weight:700;transition:background-color .3s}.notebook-link:hover{background-color:#357ae8}"""
    with open(OUTPUT_CSS_FILE, 'w', encoding='utf-8') as f: f.write(css_content)
    print(f"✅ Created '{OUTPUT_CSS_FILE}'")

def write_js_file(data):
    js_template = """
const subjectData = {data_placeholder};
function toggleStatus(e){const t=e.dataset.subject,c=parseInt(e.dataset.sectionIndex,10),a=parseInt(e.dataset.itemIndex,10),n=subjectData[t].data[c].items[a];n.status.toLowerCase().includes("done")?n.status="Pending":n.status="Done",e.textContent=n.status,e.className=n.status.toLowerCase().includes("done")?"status-done":"status-pending",updateProgress()}document.addEventListener("DOMContentLoaded",()=>{const e=document.querySelectorAll(".tab-btn"),t=document.querySelectorAll(".tab-content");e.forEach(c=>{c.addEventListener("click",()=>{e.forEach(e=>e.classList.remove("active")),t.forEach(e=>e.classList.remove("active")),c.classList.add("active"),document.getElementById(c.dataset.tab).classList.add("active")})}),renderTables(),updateProgress()});function renderTables(){for(const e in subjectData){const t=subjectData[e],c=document.getElementById(e);if(c){let a=`<h2>${t.title}</h2><div class="subject-table-container">`;t.data.forEach((e,t)=>{a+="<div>",a+=`<table><thead><tr><th colspan="2">${e.header}</th></tr></thead><tbody>`,e.items.forEach((e,n)=>{const o=e.status.toLowerCase().includes("done")?"status-done":"status-pending",d=e.hyperlink?`<a href="${e.hyperlink}" target="_blank">${e.desc}</a>`:e.desc;a+=`<tr><td>${d}</td><td class="${o}" data-subject="${c.id}" data-section-index="${t}" data-item-index="${n}" onclick="toggleStatus(this)">${e.status}</td></tr>`}),a+="</tbody></table></div>"}),a+="</div>",c.innerHTML=a}}}function updateProgress(){let e=0,t=0;for(const c in subjectData){const a=subjectData[c];let n=0,o=0;a.data.forEach(e=>{e.items.forEach(e=>{o++,e.status.toLowerCase().includes("done")&&n++})}),e+=n,t+=o;const d=o>0?n/o*100:0,s=document.getElementById(`${c}-progress-card`);s&&(s.style.borderLeftColor=a.color,s.innerHTML=`<h3>${a.title}</h3><div class="progress-container"><div class="progress-bar" style="width:${d.toFixed(1)}%;background-color:${a.color};"></div><span class="progress-label">${n} / ${o} (${d.toFixed(1)}%)</span></div>`)}const c=t>0?e/t*100:0;document.getElementById("total-progress-bar").style.width=`${c.toFixed(1)}%`,document.getElementById("total-progress-label").textContent=`${e} / ${t} (${c.toFixed(1)}%)`}
"""
    json_data = json.dumps(data, indent=4)
    final_js_content = js_template.replace("{data_placeholder}", json_data)

    with open(OUTPUT_SCRIPT_FILE, 'w', encoding='utf-8') as f:
        f.write(final_js_content)
    print(f"✅ Created interactive '{OUTPUT_SCRIPT_FILE}' with your data.")


if __name__ == "__main__":
    print("--- Starting Interactive Webpage Generator ---")
    web_data = generate_web_data()
    if web_data:
        write_html_file()
        write_css_file()
        write_js_file(web_data)
        print("\n🎉 Success! Your interactive webpage files have been created.")
        print("Upload index.html, style.css, script.js, and the 'downloaded_images' folder to GitHub.")
    else:
        print("\n❌ Webpage generation failed. Please check for errors above.")
