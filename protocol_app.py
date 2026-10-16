"""Local redline-to-protocol helper. Uses only Python's standard library."""

from __future__ import annotations

import difflib
import html
import io
import json
import os
import re
import threading
import webbrowser
import zipfile
from copy import deepcopy
from datetime import datetime
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from urllib.parse import quote, unquote, urlparse
import xml.etree.ElementTree as ET


HERE = Path(__file__).resolve().parent
TEMPLATE = HERE / "шаблон_протокола_разногласий.docx"
HOST = "127.0.0.1"
PORT = 8765
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
ET.register_namespace("w", W)
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
SESSION = {"filename": "", "document": None, "edits": []}


def tag(local: str) -> str:
    return f"{{{W}}}{local}"


def node_text(node: ET.Element, view: str) -> str:
    """Extract Word paragraph text as it appears before or after tracked changes."""
    pieces: list[str] = []

    def walk(el: ET.Element, visible: bool = True) -> None:
        local = el.tag.rsplit("}", 1)[-1]
        if local in ("ins", "moveTo") and view == "before":
            visible = False
        elif local in ("del", "moveFrom") and view == "after":
            visible = False
        if local in ("t", "delText") and visible:
            pieces.append(el.text or "")
        elif local == "tab" and visible:
            pieces.append("\t")
        elif local in ("br", "cr") and visible:
            pieces.append("\n")
        for child in el:
            walk(child, visible)

    walk(node)
    return "".join(pieces).strip()


def has_revision(node: ET.Element) -> bool:
    return any(el.tag in (tag("ins"), tag("del"), tag("moveFrom"), tag("moveTo")) for el in node.iter())


NUMBER_RE = re.compile(r"^\s*(?:(?:пункт|п\.)\s*)?(\d+(?:\.\d+){0,6})(?=\.|\s*(?:—|–|-)|\s*$)", re.I)


def clause_number(text: str) -> str | None:
    match = NUMBER_RE.match(text or "")
    return match.group(1) if match else None


def plain_text(root: ET.Element) -> str:
    return " ".join(node_text(p, "after") for p in root.findall(".//w:body/w:p", NS) if node_text(p, "after"))


def guess_fields(filename: str, root: ET.Element) -> dict[str, str]:
    paras = [node_text(p, "after") for p in root.findall(".//w:body/w:p", NS)]
    paras = [p for p in paras if p]
    full = " ".join(paras[:16])
    title = Path(filename).stem
    for p in paras[:8]:
        if len(p) > 10 and re.match(r"(?i)^договор\b", p):
            title = p.strip()
            break
    city = ""
    m = re.search(r"\bг\.\s*([А-ЯЁA-Z][А-ЯЁа-яёA-Za-z-]+)", full)
    if m:
        city = m.group(1)
    date = ""
    m = re.search(r"(?:«\s*[^»]+\s*»\s*)?(?:\d{1,2}[./]\d{1,2}[./]\d{2,4}|[А-ЯЁа-яё]{3,15}\s+20\d{2})\s*г(?:ода)?\.?", full, re.I)
    if m:
        date = m.group(0).replace("«", "").replace("»", "").strip(" \"")
    parties = {}
    for label in ("Сторона-1", "Сторона-2"):
        for p in paras[:30]:
            if re.search(r"именуем\w*\s+далее\s+[«\"]" + re.escape(label), p, re.I):
                parties[label] = re.sub(r"^\(\d+\)\s*", "", re.split(r",?\s*именуем\w*\s+далее", p, flags=re.I)[0]).strip(" ,;—-")
                break
    return {
        "title": title,
        "date": date,
        "city": city,
        "party1": parties.get("Сторона-1", "Сторона-1"),
        "party2": parties.get("Сторона-2", "Сторона-2"),
    }


def extract_edits(data: bytes) -> tuple[ET.Element, list[dict], dict[str, str]]:
    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    edits = []
    current_before = None
    current_after = None
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("В документе не найден текст договора.")
    idx = 0
    for paragraph in body.findall(".//w:p", NS):
        # Don't accidentally treat text inside tables such as signatures as clauses.
        before = node_text(paragraph, "before")
        after = node_text(paragraph, "after")
        old_number, new_number = clause_number(before), clause_number(after)
        if old_number:
            current_before = old_number
        if new_number:
            current_after = new_number
        if not has_revision(paragraph) or before == after:
            continue
        if before and after:
            label_old = old_number or current_before
            label_new = new_number or current_after
            label = f"{label_old} → {label_new}" if old_number and new_number and old_number != new_number else (label_old or label_new or "Без номера")
        elif before:
            label = old_number or current_before or "Без номера"
        else:
            label = new_number or "Без номера"
        edits.append({
            "id": idx,
            "number": label,
            "before": before,
            "after": after,
            "included": True,
            "kind": "insert" if not before and after else "delete" if before and not after else "change",
        })
        idx += 1
    return root, edits, guess_fields("Договор", root)


def safe_xml_text(text: str) -> str:
    return text


def paragraph_text(p: ET.Element) -> str:
    return "".join((t.text or "") for t in p.findall(".//w:t", NS))


def replace_in_paragraph(p: ET.Element, replacements: dict[str, str]) -> None:
    texts = p.findall(".//w:t", NS)
    if not texts:
        return
    combined = "".join(t.text or "" for t in texts)
    replaced = combined
    for old, new in replacements.items():
        replaced = re.sub(re.escape(old), lambda _m: new, replaced, flags=re.I)
    if replaced == combined:
        return
    texts[0].text = replaced
    for t in texts[1:]:
        t.text = ""


def token_diff(text: str, other: str, side: str) -> list[tuple[str, bool]]:
    # Word-level diff with whitespace retained so the generated table stays readable.
    a = re.findall(r"\s+|[^\s]+", text or "")
    b = re.findall(r"\s+|[^\s]+", other or "")
    matcher = difflib.SequenceMatcher(a=a, b=b, autojunk=False)
    ranges = []
    for op, a1, a2, b1, b2 in matcher.get_opcodes():
        if side == "before" and op in ("delete", "replace"):
            ranges.append((a1, a2, True))
        elif side == "after" and op in ("insert", "replace"):
            ranges.append((b1, b2, True))
        elif op == "equal":
            ranges.append((a1 if side == "before" else b1, a2 if side == "before" else b2, False))
    tokens = a if side == "before" else b
    spans = []
    for start, end, bold in ranges:
        if end > start:
            spans.append(("".join(tokens[start:end]), bold))
    return spans


def set_cell_runs(cell: ET.Element, text: str, compare_to: str, side: str) -> None:
    p = cell.find("w:p", NS)
    if p is None:
        p = ET.SubElement(cell, tag("p"))
    ppr = p.find("w:pPr", NS)
    for child in list(p):
        if child is not ppr:
            p.remove(child)
    spans = token_diff(text, compare_to, side) if compare_to is not None else [(text, False)]
    if not spans and text:
        spans = [(text, False)]
    for piece, bold in spans:
        if not piece:
            continue
        run = ET.SubElement(p, tag("r"))
        rpr = ET.SubElement(run, tag("rPr"))
        if bold:
            ET.SubElement(rpr, tag("b"))
        t = ET.SubElement(run, tag("t"))
        if piece[:1].isspace() or piece[-1:].isspace():
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = safe_xml_text(piece)


def set_cell_text(cell: ET.Element, text: str) -> None:
    set_cell_runs(cell, text, "", "after")


def create_protocol(selected: list[dict], fields: dict[str, str]) -> bytes:
    if not TEMPLATE.exists():
        raise ValueError("Не найден шаблон протокола рядом с программой.")
    with zipfile.ZipFile(TEMPLATE, "r") as source:
        parts = {name: source.read(name) for name in source.namelist()}
    root = ET.fromstring(parts["word/document.xml"])
    replacements = {
        "{{название договора}}": fields.get("title", ""),
        "{{дата договора}}": fields.get("date", ""),
        "{{город}}": fields.get("city", ""),
        "{{Сторона-1}}": fields.get("party1", ""),
        "{{Сторона-2}}": fields.get("party2", ""),
    }
    for p in root.findall(".//w:p", NS):
        replace_in_paragraph(p, replacements)
    tables = root.findall(".//w:tbl", NS)
    if len(tables) < 2:
        raise ValueError("В шаблоне не найдена таблица редакций сторон.")
    table = tables[1]
    rows = table.findall("w:tr", NS)
    if len(rows) < 2:
        raise ValueError("В таблице шаблона не найдена строка для правки.")
    template_row = rows[1]
    cells = template_row.findall("w:tc", NS)
    if len(cells) < 3:
        raise ValueError("В таблице шаблона должно быть три колонки.")
    insert_at = list(table).index(template_row)
    row_template = deepcopy(template_row)
    table.remove(template_row)
    for index, edit in enumerate(selected):
        row = deepcopy(row_template)
        row_cells = row.findall("w:tc", NS)
        set_cell_text(row_cells[0], edit.get("number", ""))
        before, after = edit.get("before", ""), edit.get("after", "")
        before_display = before if before else "Исключить"
        after_display = after if after else "Исключить"
        set_cell_runs(row_cells[1], before_display, after_display, "before")
        set_cell_runs(row_cells[2], after_display, before_display, "after")
        table.insert(insert_at + index, row)
    parts["word/document.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as dest:
        for name, contents in parts.items():
            dest.writestr(name, contents)
    return output.getvalue()


PAGE = r"""<!doctype html>
<html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Протокол разногласий</title><style>
:root{font-family:Segoe UI,Arial,sans-serif;color:#20252b;background:#f4f6f8}*{box-sizing:border-box}body{margin:0}
header{background:#17324d;color:#fff;padding:20px 26px}header h1{font-size:22px;margin:0 0 5px}header p{margin:0;color:#d4e0ea;font-size:14px}
main{max-width:1120px;margin:22px auto;padding:0 18px}.card{background:#fff;border:1px solid #dce2e8;border-radius:10px;padding:18px;margin-bottom:16px;box-shadow:0 2px 8px #182b3b0a}
h2{font-size:17px;margin:0 0 13px}.hint{font-size:13px;color:#64717c}.drop{border:2px dashed #9baab7;border-radius:9px;padding:25px;text-align:center;color:#526170;background:#fbfcfd;cursor:pointer}.drop.over{border-color:#2769a8;background:#eff7ff}.drop strong{display:block;color:#20364a;margin-bottom:5px}.drop input{display:none}
.grid{display:grid;grid-template-columns:repeat(2,minmax(200px,1fr));gap:12px}.field label{display:block;font-size:12px;color:#596671;margin-bottom:5px}.field input{width:100%;padding:9px 10px;border:1px solid #cad3dc;border-radius:6px;font:inherit}.edit{border-top:1px solid #e5e9ed;padding:14px 0;display:grid;grid-template-columns:30px 100px 1fr 1fr;gap:12px;align-items:start}.edit:first-child{border-top:0}.number{font-weight:650;color:#24394c;padding-top:3px;overflow-wrap:anywhere}.side{font-size:11px;text-transform:uppercase;color:#75808b;margin-bottom:4px}.text{font-size:13px;line-height:1.5;white-space:pre-wrap;overflow-wrap:anywhere}.del{background:#ffe0e0;color:#8f2828;text-decoration:line-through}.ins{background:#d9f3df;color:#155b2a}.empty{color:#7a8590;font-style:italic}.btn{border:0;border-radius:6px;padding:10px 15px;font:inherit;font-weight:600;cursor:pointer}.primary{background:#1769aa;color:#fff}.primary:disabled{opacity:.5;cursor:wait}.secondary{background:#e7edf2;color:#263f52}.topline{display:flex;gap:12px;align-items:center;justify-content:space-between;margin-bottom:12px}.msg{font-size:13px;color:#44535f}.counter{font-size:13px;color:#586774}.warning{color:#a24015;background:#fff4e9;padding:10px;border-radius:6px;font-size:13px;margin-bottom:12px}.hidden{display:none}@media(max-width:720px){.edit{grid-template-columns:26px 82px 1fr}.edit .after{grid-column:3}.grid{grid-template-columns:1fr}}
</style></head><body>
<header><h1>Протокол разногласий</h1><p>Обработка DOCX выполняется на этом компьютере. Файлы не отправляются в интернет.</p></header>
<main><section class="card"><h2>1. Загрузите договор с исправлениями</h2><label id="drop" class="drop"><strong>Перетащите файл сюда</strong>или нажмите, чтобы выбрать DOCX через проводник<input id="file" type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"></label><div id="msg" class="msg" style="margin-top:10px"></div></section>
<section id="fieldsCard" class="card hidden"><h2>2. Данные протокола</h2><div class="grid"><div class="field"><label>Название договора</label><input id="title"></div><div class="field"><label>Дата договора</label><input id="date"></div><div class="field"><label>Город</label><input id="city"></div><div class="field"><label>Сторона-1</label><input id="party1"></div><div class="field"><label>Сторона-2</label><input id="party2"></div></div></section>
<section id="editsCard" class="card hidden"><div class="topline"><h2 style="margin:0">3. Проверьте пункты и отметьте нужные</h2><div id="counter" class="counter"></div></div><div id="edits"></div><div style="display:flex;gap:10px;align-items:center;margin-top:15px"><button id="build" class="btn primary">Создать протокол по шаблону</button><button id="all" class="btn secondary">Выбрать все</button><span id="buildMsg" class="msg"></span></div></section></main>
<script>
const $=id=>document.getElementById(id);let edits=[];const file=$('file'),drop=$('drop');
function esc(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function diff(a,b,side){const x=(a.match(/\s+|[^\s]+/g)||[]),y=(b.match(/\s+|[^\s]+/g)||[]);let out='';let i=0,j=0;const n=x.length,m=y.length;const dp=Array.from({length:n+1},()=>Array(m+1).fill(0));for(i=n-1;i>=0;i--)for(j=m-1;j>=0;j--)dp[i][j]=x[i]===y[j]?1+dp[i+1][j+1]:Math.max(dp[i+1][j],dp[i][j+1]);i=j=0;let buf='',kind='';function flush(){if(buf)out+=kind==='same'?esc(buf):`<span class="${kind}">${esc(buf)}</span>`;buf=''}while(i<n||j<m){let k,v;if(i<n&&j<m&&x[i]===y[j]){k='same';v=x[i];i++;j++}else if(i<n&&(j===m||dp[i+1][j]>=dp[i][j+1])){k=side==='before'?'del':'skip';v=x[i++]}else{k=side==='after'?'ins':'skip';v=y[j++]}if(k==='skip')continue;if(kind!==k)flush();kind=k;buf+=v}flush();return out||'<span class="empty">—</span>'}
function render(){const box=$('edits');box.innerHTML='';edits.forEach((e,i)=>{const row=document.createElement('div');row.className='edit';row.innerHTML=`<input type="checkbox" ${e.included?'checked':''} aria-label="Включить пункт"><div class="number">${esc(e.number)}</div><div><div class="side">Было</div><div class="text">${e.before?diff(e.before,e.after,'before'):'<span class="empty">Отсутствует</span>'}</div></div><div class="after"><div class="side">Стало</div><div class="text">${e.after?diff(e.before,e.after,'after'):'<span class="empty">Исключить</span>'}</div></div>`;row.querySelector('input').onchange=ev=>{e.included=ev.target.checked;count()};box.append(row)});$('fieldsCard').classList.remove('hidden');$('editsCard').classList.remove('hidden');count()}
function count(){$('counter').textContent=`Отмечено: ${edits.filter(e=>e.included).length} из ${edits.length}`}
async function loadFile(f){if(!f)return;if(!f.name.toLowerCase().endsWith('.docx')){$('msg').textContent='Нужен файл Word с расширением .docx.';return}$('msg').textContent=`Читаю локально: ${f.name}…`;$('buildMsg').textContent='';try{const res=await fetch('/api/parse?name='+encodeURIComponent(f.name),{method:'POST',body:f,headers:{'Content-Type':'application/vnd.openxmlformats-officedocument.wordprocessingml.document'}});const data=await res.json();if(!res.ok)throw Error(data.error||'Не удалось прочитать договор');edits=data.edits;$('msg').textContent=`Загружен ${f.name}. Найдено правок: ${edits.length}.`;for(const [k,v] of Object.entries(data.fields))$(k).value=v;if(!edits.length)$('msg').textContent+=' В документе не найдено отслеживаемых исправлений Word.';render()}catch(e){$('msg').textContent=e.message}}
file.addEventListener('change',()=>loadFile(file.files[0]));drop.addEventListener('dragover',e=>{e.preventDefault();drop.classList.add('over')});drop.addEventListener('dragleave',()=>drop.classList.remove('over'));drop.addEventListener('drop',e=>{e.preventDefault();drop.classList.remove('over');loadFile(e.dataTransfer.files[0])});
$('all').onclick=()=>{const val=edits.some(e=>!e.included);edits.forEach(e=>e.included=val);render()};
$('build').onclick=async()=>{const btn=$('build');btn.disabled=true;$('buildMsg').textContent='Создаю документ…';try{const payload={edits:edits.filter(e=>e.included),fields:{title:$('title').value,date:$('date').value,city:$('city').value,party1:$('party1').value,party2:$('party2').value}};if(!payload.edits.length)throw Error('Отметьте хотя бы один пункт.');const res=await fetch('/api/build',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});if(!res.ok){const err=await res.json();throw Error(err.error)}const blob=await res.blob(),a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='Протокол разногласий.docx';a.click();URL.revokeObjectURL(a.href);$('buildMsg').textContent='Готово. Документ сохранён в папку загрузок браузера.'}catch(e){$('buildMsg').textContent=e.message}finally{btn.disabled=false}};
</script></body></html>"""


class Handler(BaseHTTPRequestHandler):
    def log_message(self, _format: str, *args) -> None:
        # No file names or document contents are written to logs.
        return

    def send_json(self, status: int, payload: dict) -> None:
        raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(raw)

    def do_GET(self) -> None:
        if urlparse(self.path).path != "/":
            self.send_error(404)
            return
        raw = PAGE.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(raw)

    def do_POST(self) -> None:
        path = urlparse(self.path)
        length = int(self.headers.get("Content-Length", "0"))
        if length > 40 * 1024 * 1024:
            self.send_json(413, {"error": "Файл больше 40 МБ."})
            return
        body = self.rfile.read(length)
        if path.path == "/api/parse":
            try:
                if not zipfile.is_zipfile(io.BytesIO(body)):
                    raise ValueError("Файл не похож на DOCX. Сохраните договор в формате .docx.")
                root, edits, _ = extract_edits(body)
                SESSION.update(filename=unquote(path.query.split("name=", 1)[-1]) if "name=" in path.query else "Договор.docx", document=body, edits=edits)
                fields = guess_fields(SESSION["filename"], root)
                self.send_json(200, {"edits": edits, "fields": fields})
            except (ValueError, KeyError, zipfile.BadZipFile, ET.ParseError) as exc:
                self.send_json(400, {"error": str(exc)})
            return
        if path.path == "/api/build":
            try:
                payload = json.loads(body.decode("utf-8"))
                selected = payload.get("edits") or []
                fields = payload.get("fields") or {}
                output = create_protocol(selected, fields)
                filename = "Протокол разногласий.docx"
                self.send_response(200)
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                self.send_header("Content-Disposition", "attachment; filename*=UTF-8''" + quote(filename))
                self.send_header("Content-Length", str(len(output)))
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                self.wfile.write(output)
            except (ValueError, KeyError, zipfile.BadZipFile, ET.ParseError, json.JSONDecodeError) as exc:
                self.send_json(400, {"error": str(exc)})
            return
        self.send_error(404)


def main() -> None:
    if not TEMPLATE.exists():
        raise SystemExit(f"Рядом с программой не найден шаблон: {TEMPLATE.name}")
    server = HTTPServer((HOST, PORT), Handler)
    url = f"http://{HOST}:{PORT}"
    print(f"Программа работает локально: {url}")
    print("Закройте это окно, чтобы остановить программу.")
    threading.Timer(0.7, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
