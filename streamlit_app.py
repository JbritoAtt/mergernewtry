import os
import io
import re
import json
import zipfile
import tempfile
from datetime import datetime
from pathlib import Path

import streamlit as st
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docxtpl import DocxTemplate
from jinja2 import Environment

# -------------------------------
# Constants & Helpers
# -------------------------------
APP_DIR = Path.cwd()
TEMPLATE_DIR = APP_DIR / "templates_store"
OUTPUT_DIR = APP_DIR / "_generated"
TEMPLATE_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

FIELD_REGEX = re.compile(r"{{\s*([a-zA-Z0-9_\.]+)\s*}}")
FOR_REGEX = re.compile(r"{%\s*for\s+(\w+)\s+in\s+([a-zA-Z0-9_\.]+)\s*%}")


def slugify(name: str) -> str:
    return re.sub(r"[^a-zA-Z0-9_-]+", "-", name.strip()).strip("-").lower()


def _add_page_field(paragraph, prefix=None, suffix=None):
    if prefix:
        paragraph.add_run(prefix)
    r = paragraph.add_run()
    begin = OxmlElement('w:fldChar'); begin.set(qn('w:fldCharType'), 'begin'); r._element.append(begin)
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve'); instr.text = ' PAGE '
    r._element.append(instr)
    sep = OxmlElement('w:fldChar'); sep.set(qn('w:fldCharType'), 'separate'); r._element.append(sep)
    txt = OxmlElement('w:t'); txt.text = '1'; r._element.append(txt)
    end = OxmlElement('w:fldChar'); end.set(qn('w:fldCharType'), 'end'); r._element.append(end)
    if suffix:
        paragraph.add_run(suffix)


# -------------------------------
# Template creation (Builder) -> .docx with Jinja placeholders
# -------------------------------

def build_docx_from_blocks(blocks: list, base_docx: bytes | None = None, default_font=("Calibri", 11)) -> bytes:
    doc = Document(io.BytesIO(base_docx)) if base_docx else Document()
    try:
        normal = doc.styles["Normal"]
        normal.font.name = default_font[0]
        normal.font.size = Pt(default_font[1])
    except Exception:
        pass

    # Footer page number centered
    section = doc.sections[0]
    footer_p = section.footer.paragraphs[0] if section.footer.paragraphs else section.footer.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_page_field(footer_p)

    for b in blocks:
        t = b.get("type")
        if t == "heading":
            doc.add_heading(b.get("text", ""), level=int(b.get("level", 1)))
        elif t == "paragraph":
            p = doc.add_paragraph(b.get("text", ""))
        elif t == "page_break":
            doc.add_page_break()
        elif t == "raw":
            # For power users: insert literal Jinja tags, e.g. {% for item in items %} ... {% endfor %}
            doc.add_paragraph(b.get("text", ""))
        elif t == "table":
            # simple table from headers, with a for-loop row placeholder
            headers = [h.strip() for h in b.get("headers", []) if str(h).strip()]
            iterator = b.get("iterator", "row")
            source = b.get("source", "items")
            if headers:
                table = doc.add_table(rows=1, cols=len(headers))
                hdr = table.rows[0].cells
                for i, col in enumerate(headers):
                    hdr[i].text = str(col)
                # jinja loop rows as raw text in new paragraphs under the table
                doc.add_paragraph(f"{{% for {iterator} in {source} %}}")  # note the doubled braces around the Jinja tag
                row_line = " | ".join([f"{{{{ {iterator}.{slugify(h)} }}}}" for h in headers]) 
                doc.add_paragraph(row_line)
                doc.add_paragraph("{% endfor %}")
        else:
            doc.add_paragraph(b.get("text", ""))

    bio = io.BytesIO(); doc.save(bio); bio.seek(0)
    return bio.read()


# -------------------------------
# Jinja environment with helpful filters (dates, currency, upper)
# -------------------------------

def jinja_env():
    env = Environment(autoescape=False)
    def datefmt(value, fmt="%-d %B %Y"):
        if value is None:
            return ""
        if isinstance(value, (int, float)):
            dt = datetime.fromtimestamp(value)
        elif isinstance(value, datetime):
            dt = value
        else:
            try:
                dt = pd.to_datetime(value)
            except Exception:
                return str(value)
        try:
            return dt.strftime(fmt)
        except Exception:
            return str(dt)
    def gbp(value):
        try:
            return f"£{float(value):,.2f}"
        except Exception:
            return str(value)
    env.filters['datefmt'] = datefmt
    env.filters['gbp'] = gbp
    env.filters['upper'] = lambda x: str(x).upper()
    return env


# -------------------------------
# Storage helpers
# -------------------------------

def save_template(name: str, docx_bytes: bytes, meta: dict) -> Path:
    slug = slugify(name)
    folder = TEMPLATE_DIR / slug
    folder.mkdir(exist_ok=True)
    (folder / "meta.json").write_text(json.dumps(meta, indent=2), encoding="utf-8")
    (folder / f"{slug}.docx").write_bytes(docx_bytes)
    return folder


def list_templates():
    entries = []
    for p in TEMPLATE_DIR.iterdir():
        if p.is_dir() and (p / "meta.json").exists():
            try:
                meta = json.loads((p / "meta.json").read_text("utf-8"))
                entries.append({"name": meta.get("name", p.name), "slug": p.name, "path": p})
            except Exception:
                pass
    return sorted(entries, key=lambda x: x["name"].lower())


def read_template(slug: str):
    folder = TEMPLATE_DIR / slug
    meta = json.loads((folder / "meta.json").read_text("utf-8"))
    docx_path = next(folder.glob("*.docx"))
    return meta, docx_path


# -------------------------------
# Introspection helpers
# -------------------------------

def extract_vars_from_docx(docx_path: Path):
    # Best-effort: read XML and regex for {{ var }} and for-loops
    import zipfile
    vars_found = set()
    loops = []
    with zipfile.ZipFile(docx_path, 'r') as z:
        for name in z.namelist():
            if name.endswith(".xml"):
                try:
                    xml = z.read(name).decode("utf-8", errors="ignore")
                except Exception:
                    continue
                for m in FIELD_REGEX.finditer(xml):
                    vars_found.add(m.group(1))
                for m in FOR_REGEX.finditer(xml):
                    loops.append({"iter": m.group(1), "source": m.group(2)})
    return sorted(vars_found), loops


# -------------------------------
# Rendering (mail merge)
# -------------------------------

def render_one(docx_path: Path, context: dict, out_path: Path):
    tpl = DocxTemplate(str(docx_path))
    env = jinja_env()
    tpl.render(context, jinja_env=env)
    tpl.save(str(out_path))


# -------------------------------
# Streamlit UI
# -------------------------------
st.set_page_config(page_title="Template Studio + Mail Merge", layout="wide")
st.title("📄 Template Studio + Mail Merge (MVP)")

page = st.sidebar.radio("Navigate", ["Templates", "Data", "Generate", "Help"], index=0)

if page == "Templates":
    st.subheader("Create or Upload Templates")
    tabs = st.tabs(["Builder (no Word)", "Upload .docx (with {{placeholders}})", "Manage"])    

    with tabs[0]:
        st.caption("Design a simple letter using blocks. You can use placeholders like {{ full_name }} anywhere.")
        if "builder_blocks" not in st.session_state:
            st.session_state.builder_blocks = []
        colA, colB = st.columns([2, 1])
        with colA:
            st.write("### Blocks")
            for i, b in enumerate(st.session_state.builder_blocks):
                with st.expander(f"Block {i+1}: {b.get('type')}"):
                    if b.get('type') in ("heading", "paragraph", "raw"):
                        b['text'] = st.text_area("Text", b.get('text', ''), key=f"txt_{i}")
                    if b.get('type') == "heading":
                        b['level'] = st.selectbox("Level", [1,2,3,4], index=int(b.get('level',1))-1, key=f"lvl_{i}")
                    if b.get('type') == "table":
                        b['headers'] = st.text_input("Headers (comma separated)", ", ".join(b.get('headers', [])), key=f"hdr_{i}")
                        b['iterator'] = st.text_input("Iterator name", b.get('iterator','row'), key=f"itr_{i}")
                        b['source'] = st.text_input("List source variable", b.get('source','items'), key=f"src_{i}")
                    if st.button("Delete", key=f"del_{i}"):
                        st.session_state.builder_blocks.pop(i)
                        st.experimental_rerun()
        with colB:
            st.write("### Add Block")
            bt = st.selectbox("Type", ["heading", "paragraph", "table", "page_break", "raw"])
            if st.button("➕ Add"):
                block = {"type": bt, "text": ""}
                if bt == "heading":
                    block["level"] = 1
                    block["text"] = "Heading"
                if bt == "paragraph":
                    block["text"] = "Enter paragraph text with placeholders like {{ full_name }}."
                if bt == "raw":
                    block["text"] = "{% if show_extra %}Extra text here{% endif %}"
                if bt == "table":
                    block["headers"] = ["Item", "Qty"]
                    block["iterator"] = "row"
                    block["source"] = "items"
                st.session_state.builder_blocks.append(block)
                st.experimental_rerun()

        st.write("---")
        name = st.text_input("Template name", "Generic Letter")
        base = st.file_uploader("Optional: base .docx with your letterhead/styles", type=["docx"])
        if st.button("Save Template"):
            blocks = st.session_state.builder_blocks
            base_bytes = base.read() if base else None
            docx_bytes = build_docx_from_blocks(blocks, base_bytes)
            meta = {"name": name, "created": datetime.utcnow().isoformat(), "source": "builder"}
            folder = save_template(name, docx_bytes, meta)
            st.success(f"Saved as {folder.name}")

    with tabs[1]:
        st.caption("Upload a .docx that already contains Jinja placeholders like {{ full_name }} and loops.")
        up = st.file_uploader("Upload .docx template", type=["docx"], key="updocx")
        tname = st.text_input("Name for this template", "Uploaded Template")
        if up and st.button("Save uploaded template"):
            docx_bytes = up.read()
            meta = {"name": tname, "created": datetime.utcnow().isoformat(), "source": "upload"}
            folder = save_template(tname, docx_bytes, meta)
            st.success(f"Saved as {folder.name}")

    with tabs[2]:
        st.caption("Inspect existing templates and see required variables.")
        rows = list_templates()
        if not rows:
            st.info("No templates yet.")
        else:
            for t in rows:
                meta, path = read_template(t['slug'])
                vars_found, loops = extract_vars_from_docx(path)
                with st.expander(f"{meta.get('name')} ({t['slug']})"):
                    st.write(f"Path: {path.name}")
                    st.write("**Variables found:**", vars_found if vars_found else "(none)")
                    if loops:
                        st.write("**Loops:**")
                        for L in loops:
                            st.code(f"for {L['iter']} in {L['source']}")

elif page == "Data":
    st.subheader("Upload Data for Mail Merge")
    st.caption("CSV or Excel. One row = one document. For repeating sections, include a JSON column (e.g., 'items') that contains a list of dicts.")
    up = st.file_uploader("Upload CSV/XLSX", type=["csv", "xlsx"]) 
    if up:
        if up.name.lower().endswith(".csv"):
            df = pd.read_csv(up)
        else:
            df = pd.read_excel(up)
        st.session_state['data_df'] = df
        st.dataframe(df.head(50))
        st.success(f"Loaded {len(df)} rows.")

        # Optional quick JSON column prettifier
        json_cols = [c for c in df.columns if df[c].astype(str).str.startswith('[').any() or df[c].astype(str).str.startswith('{').any()]
        if json_cols:
            st.caption("Detected potential JSON columns for loops:")
            st.code(", ".join(json_cols))

else:
    if page == "Generate":
        st.subheader("Generate Documents")
        templates = list_templates()
        if not templates:
            st.info("Create or upload a template first.")
        elif 'data_df' not in st.session_state:
            st.info("Upload data in the Data tab.")
        else:
            tnames = [t['name'] for t in templates]
            choice = st.selectbox("Choose template", tnames)
            t = next(x for x in templates if x['name'] == choice)
            meta, docx_path = read_template(t['slug'])
            vars_found, _ = extract_vars_from_docx(docx_path)

            df = st.session_state['data_df']
            st.write("**Columns in your data:**", list(df.columns))
            st.write("**Template variables:**", vars_found)

            name_pattern = st.text_input("Output filename pattern", "{{full_name}}_Letter.docx")
            start = st.number_input("Start row (1-based)", min_value=1, value=1, step=1)
            end = st.number_input("End row (inclusive)", min_value=1, value=len(df), step=1)
            end = min(end, len(df))

            # Map JSON columns (for loops)
            st.write("### JSON/List columns (optional)")
            json_map = {}
            for col in df.columns:
                sample = df.iloc[0][col]
                if isinstance(sample, str) and (sample.strip().startswith("[") or sample.strip().startswith("{")):
                    json_map[col] = st.text_input(f"Use column '{col}' as variable (e.g., items)", value="")

            if st.button("Generate ZIP"):
                rows = df.iloc[start-1:end].to_dict(orient="records")
                tmpdir = tempfile.mkdtemp(prefix="gen_")
                out_paths = []
                for i, row in enumerate(rows, start=start):
                    ctx = {}
                    # copy row values
                    for k, v in row.items():
                        # try to parse JSON columns if mapped
                        v_out = v
                        if isinstance(v, str) and k in json_map and json_map[k].strip():
                            try:
                                v_out = json.loads(v)
                            except Exception:
                                v_out = v
                            ctx[json_map[k].strip()] = v_out
                        ctx[slugify(str(k))] = v  # also a slugified version for convenience
                        ctx[k] = v
                    # filename
                    try:
                        env = jinja_env()
                        fname = Environment().from_string(name_pattern).render(**ctx)
                    except Exception:
                        fname = f"doc_{i}.docx"
                    if not fname.lower().endswith('.docx'):
                        fname += '.docx'
                    out_path = Path(tmpdir) / fname
                    out_path.parent.mkdir(parents=True, exist_ok=True)
                    try:
                        render_one(docx_path, ctx, out_path)
                        out_paths.append(out_path)
                    except Exception as e:
                        error_path = Path(tmpdir) / f"error_row_{i}.txt"
                        error_path.write_text(str(e))
                        out_paths.append(error_path)

                # zip them
                zip_bytes = io.BytesIO()
                with zipfile.ZipFile(zip_bytes, 'w', zipfile.ZIP_DEFLATED) as z:
                    for p in out_paths:
                        z.write(p, arcname=p.name)
                zip_bytes.seek(0)
                st.download_button("Download ZIP", zip_bytes, file_name="letters.zip")
                st.success(f"Generated {len(out_paths)} files.")

    else:  # Help
        st.subheader("How to use")
        st.markdown(
            """
            1. **Templates → Builder**: create a letter with headings/paragraphs. Insert placeholders like `{{ full_name }}`.
               Or **Upload .docx** that already contains placeholders and Jinja loops.
            2. **Data**: upload CSV/Excel. One row per letter. For repeating sections, put a JSON list in a column
               (e.g., `items` = `[{"name":"ISA","value":12345},{"name":"SIPP","value":9876}]`).
            3. **Generate**: pick template + rows, set filename pattern (e.g., `{{full_name}}_Letter.docx`), generate ZIP.

            **Tips**
            - Use filters in placeholders (e.g., `{{ charge | gbp }}`, `{{ letter_date | datefmt('%d %b %Y') }}`, `{{ name | upper }}`).
            - For loops, add a `raw` block with `{% for item in items %} ... {{ item.name }} ... {% endfor %}`
            - Upload a branded `base .docx` in the Builder to inherit styles, margins, headers/footers.
            - Page numbers are added automatically to the footer.
            """
        )
