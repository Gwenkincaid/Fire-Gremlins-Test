import hmac
import io
import re
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


def _secret(name):
    try:
        return str(st.secrets.get(name, ""))
    except Exception:
        return ""


def _extract(files, roster):
    import pytesseract
    from PIL import Image, ImageOps
    import unicodedata
    def normalize(value):
        return ''.join(c for c in unicodedata.normalize('NFKD', value.casefold()) if c.isalnum())
    names = {}
    for name in roster:
        names.setdefault(normalize(name), []).append(name)
    rows = []
    for f in files:
        image = ImageOps.exif_transpose(Image.open(io.BytesIO(f.getvalue()))).convert('L')
        w, h = image.size
        # Relative coordinates for the supplied landscape member-list layout.
        def read(x1, x2, y1, y2, numeric=False):
            crop = image.crop((round(w*x1),round(h*y1),round(w*x2),round(h*y2)))
            crop = ImageOps.autocontrast(crop)
            crop = ImageOps.invert(crop).resize((crop.width*3,crop.height*3))
            config = '--psm 7' + (' -c tessedit_char_whitelist=0123456789,' if numeric else '')
            return pytesseract.image_to_string(crop,lang='eng',config=config).strip()
        for i in range(6):
            y1 = .385 + i*.1015
            y2 = y1 + .080
            rank = read(.149,.166,y1,y2,True)
            name = read(.239,.419,y1,y2)
            power = read(.441,.548,y1,y2,True).replace(',','').replace(' ','')
            if not rank.isdigit() or not power.isdigit() or not name:
                continue
            matches = names.get(normalize(name), [])
            if len(matches)==1:
                name=matches[0]
            rows.append({'rank':int(rank),'name':name,'power':int(power)})
    if not rows:
        raise ValueError('No complete rows recognized. Use full landscape member-list screenshots matching the original layout.')
    return rows


def _workbook_bytes(source, records, day, report_date):
    wb = load_workbook(io.BytesIO(source))
    if "Player Progress" not in wb.sheetnames:
        raise ValueError("Workbook is missing the Player Progress worksheet.")
    ws = wb["Player Progress"]
    target = None
    for c in range(2, ws.max_column + 1):
        label = str(ws.cell(1,c).value or "")
        if re.fullmatch(rf"Day\s*{day}(?:\s*\([^)]*\))?",label.strip(),re.I):
            target = c
            break
    if target is None:
        target = ws.max_column + 1
    ws.cell(1,target).value = f"Day {day} ({report_date.strftime('%m/%d/%Y')})"
    existing = {}
    for r in range(2,ws.max_row+1):
        name = ws.cell(r,1).value
        if name is not None and str(name).strip():
            existing[str(name).strip()] = r
    for item in records:
        name = str(item["Player"]).strip()
        if not name or name.casefold() == "alliance power":
            continue
        row = existing.get(name)
        if row is None:
            row = max(ws.max_row+1,2)
            ws.cell(row,1).value = name
            existing[name] = row
        ws.cell(row,target).value = int(item["Power"])
        ws.cell(row,target).number_format = '#,##0'
    ws.column_dimensions[get_column_letter(target)].width = 24
    result = io.BytesIO()
    wb.save(result)
    return result.getvalue()


def show_admin():
    st.markdown('<div class="banner"><div class="eyebrow">THE WAR COUNCIL · PRIVATE ADMIN</div>'
                '<h1>📜 Screenshot Importer</h1><div class="subtitle">'
                'Extract, review, and prepare the next chronicle.</div></div>',
                unsafe_allow_html=True)
    expected = _secret("ADMIN_PASSWORD")
    if not expected:
        st.warning("Add ADMIN_PASSWORD to Streamlit Secrets to enable the admin page.")
        return
    if not st.session_state.get("admin_authenticated"):
        with st.form("admin_login"):
            entered = st.text_input("Admin passphrase",type="password")
            submit = st.form_submit_button("Unlock Admin")
        if submit:
            if hmac.compare_digest(entered.encode(),expected.encode()):
                st.session_state.admin_authenticated = True
                st.rerun()
            st.error("Incorrect admin passphrase.")
        return
    if st.button("Lock admin"):
        st.session_state.admin_authenticated = False
        st.session_state.pop("extracted_members",None)
        st.rerun()
    st.info("Step 1: Upload screenshots. Step 2: Review every row. Step 3: Download the updated Excel workbook. "
            "Nothing is written to GitHub automatically.")
    alliance = st.selectbox("Alliance",["Fire Gremlins","NONAME","MALICE","Other"])
    other = st.text_input("Alliance name") if alliance=="Other" else ""
    name = other.strip() if alliance=="Other" else alliance
    day = st.number_input("Report day",min_value=1,max_value=999,value=6,step=1)
    report_date = st.date_input("Report date",value=date.today())
    st.caption("Upload the Excel workbook belonging to the selected alliance. "
               "The Fire Gremlins workbook can be loaded from the deployed repository by default.")
    base_path=Path(__file__).with_name("progress.xlsx")
    base = st.file_uploader("Alliance Excel workbook (optional for Fire Gremlins)",type=["xlsx"],key="base")
    source = base.getvalue() if base else (base_path.read_bytes() if alliance=="Fire Gremlins" and base_path.exists() else None)
    roster=[]
    if source:
        try:
            frame=pd.read_excel(io.BytesIO(source),sheet_name="Player Progress")
            roster=frame.iloc[:,0].dropna().astype(str).str.strip().tolist()
        except Exception as e:
            st.error(f"Cannot read the selected workbook: {e}")
            return
    else:
        st.warning("Upload this alliance's Excel workbook before saving its report.")
    files=st.file_uploader("Member-list screenshots",type=["png","jpg","jpeg"],
                           accept_multiple_files=True,key="screenshots")
    st.caption("Free local OCR • landscape screenshots • review every value; missed rows must be added manually.")
    if st.button("Extract player data",disabled=not files):
        st.session_state.pop("extracted_members", None)
        st.session_state.pop("review_editor", None)
        try:
            with st.spinner("Reading screenshots..."):
                rows=_extract(files,roster)
            by_rank={}
            conflicts=[]
            for item in rows:
                rank=int(item["rank"])
                if rank in by_rank and by_rank[rank]!=item:
                    conflicts.append(rank)
                else:
                    by_rank[rank]=item
            cleaned=[{"Rank":r,"Player":str(v["name"]),"Power":int(v["power"])}
                     for r,v in sorted(by_rank.items())]
            st.session_state.extracted_members=cleaned
            st.session_state.extract_context=(name,int(day),str(report_date))
            st.session_state.extract_conflicts=conflicts
        except Exception as e:
            st.error(f"Extraction failed: {e}")
    if "extracted_members" not in st.session_state:
        return
    if st.session_state.get("extract_context") != (name,int(day),str(report_date)):
        st.warning("Alliance, day, or date changed. Run extraction again before saving.")
        return
    conflicts=st.session_state.get("extract_conflicts",[])
    if conflicts:
        st.warning(f"Conflicting duplicate screenshot ranks: {conflicts}. Check carefully.")
    frame=pd.DataFrame(st.session_state.extracted_members)
    edited=st.data_editor(frame,num_rows="dynamic",hide_index=True,use_container_width=True,
                          column_config={"Rank":st.column_config.NumberColumn(required=True,step=1),
                                         "Player":st.column_config.TextColumn(required=True),
                                         "Power":st.column_config.NumberColumn(required=True,step=1)},
                          key="review_editor")
    if edited.empty:
        st.warning("No members extracted.")
        return
    edited["Player"]=edited["Player"].astype(str).str.strip()
    if edited["Player"].eq("").any() or edited["Player"].duplicated().any():
        st.error("Resolve blank or duplicate player names.")
        return
    if edited["Rank"].isna().any() or edited["Rank"].duplicated().any():
        st.error("Resolve missing or duplicate ranks.")
        return
    if edited["Power"].isna().any() or (edited["Power"]<=0).any():
        st.error("Every power value must be positive.")
        return
    if roster:
        unknown=sorted(set(edited["Player"])-set(roster))
        missing=sorted(set(roster)-set(edited["Player"]))
        if unknown:
            st.warning("Names not found in existing roster (check spelling): "+", ".join(unknown))
        if missing:
            st.warning(f"{len(missing)} roster members were not found in these screenshots. "
                       "Missing values will remain blank for this day.")
    st.metric("Extracted members",len(edited))
    st.metric("Sum of extracted player power",f'{int(edited["Power"].sum()):,}')
    if not source:
        return
    st.warning("Check that all expected ranks are present before downloading. OCR can miss rows or misread digits.")
    confirmed=st.checkbox("I reviewed the extracted names, ranks, and power values.")
    if confirmed:
        try:
            updated=_workbook_bytes(source,edited.to_dict("records"),int(day),report_date)
            st.download_button("⬇️ Download updated Excel workbook",updated,
                               file_name=f'{re.sub(r"[^A-Za-z0-9_-]+","_",name)}_Day_{day}_progress.xlsx',
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.caption("Replace progress.xlsx in the corresponding alliance's GitHub repository "
                       "to update its public dashboard. Downloading does not publish the data.")
        except Exception as e:
            st.error(f"Could not create the updated workbook: {e}")
