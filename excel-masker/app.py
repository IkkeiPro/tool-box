"""Local Excel masking app. Python 3.10+, standard library only."""
import base64
import io
import json
import re
import secrets
import threading
import webbrowser
import zipfile
from pathlib import Path
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from xml.dom import minidom

ROOT = Path(__file__).parent
TOKEN = secrets.token_urlsafe(32)
MAX_REQUEST = 150 * 1024 * 1024

def mapping_rules(rows):
    mapping = {}
    for row in rows:
        name, alias = str(row.get('name', '')).strip(), str(row.get('alias', '')).strip()
        if not name or not alias:
            raise ValueError('対応表の個人名と置換後の名前をすべて入力してください。')
        if name in mapping:
            raise ValueError('同じ個人名が複数行あります。フルネームなどで区別してください。')
        if any(ord(c) < 32 for c in name + alias):
            raise ValueError('名前に改行や制御文字は使えません。')
        mapping[name] = alias
    if not mapping or len(mapping) > 5000:
        raise ValueError('対応表は1〜5,000件にしてください。')
    if any(n in a for n in mapping for a in mapping.values()):
        raise ValueError('置換後の名前に、対応表の個人名が含まれています。別の名前にしてください。')
    return mapping

def mask_xlsx(data, mapping):
    pattern = re.compile('|'.join(re.escape(n) for n in sorted(mapping, key=len, reverse=True)))
    counts = {n: 0 for n in mapping}
    def replace(s):
        def sub(m):
            counts[m[0]] += 1
            return mapping[m[0]]
        return pattern.sub(sub, s)
    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as source:
        infos = source.infolist()
        if len(infos) > 10000 or sum(i.file_size for i in infos) > 300 * 1024 * 1024:
            raise ValueError('展開後のサイズが大きすぎます（上限300MB）。')
        names = [i.filename for i in infos]
        if len(set(names)) != len(names) or 'xl/workbook.xml' not in names:
            raise ValueError('有効なxlsxファイルではありません。')
        if any(n.startswith(('xl/embeddings/', 'xl/activeX/', '_xmlsignatures/')) or n.endswith('vbaProject.bin') for n in names):
            raise ValueError('埋め込みファイル・マクロ・電子署名付きブックは未対応です。')
        warnings = []
        if any(n.startswith('xl/media/') for n in names):
            warnings.append('画像内の文字は置換・検査できません。画像を確認してください。')
        blocked = set()
        with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as target:
            for info in infos:
                raw = source.read(info)
                if info.filename.endswith(('.xml', '.rels', '.vml')):
                    if b'<!DOCTYPE' in raw.upper() or b'<!ENTITY' in raw.upper():
                        raise ValueError('DTDを含むXMLは処理できません。')
                    doc = minidom.parseString(raw)
                    changed = False
                    # Join rich-text runs so names split by character formatting are found.
                    for el in list(doc.getElementsByTagName('*')):
                        if el.localName in ('si', 'is', 'text', 'p'):
                            texts = [n for n in el.getElementsByTagName('*') if n.localName == 't']
                            old = ''.join(''.join(c.data for c in n.childNodes if c.nodeType in (3, 4)) for n in texts)
                            if texts and pattern.search(old):
                                value = replace(old)
                                for index, node in enumerate(texts):
                                    for child in list(node.childNodes):
                                        node.removeChild(child)
                                    node.appendChild(doc.createTextNode(value if index == 0 else ''))
                                changed = True
                    for el in list(doc.getElementsByTagName('*')):
                        # Formula identifiers/references remain intact; mask quoted literals only.
                        for node in list(el.childNodes):
                            if node.nodeType not in (3, 4) or not pattern.search(node.data):
                                continue
                            if el.localName in ('f', 'formula', 'formula1', 'formula2', 'definedName'):
                                new = re.sub(r'"(?:[^"]|"")*"', lambda m: '"' + replace(m[0][1:-1].replace('""', '"')).replace('"', '""') + '"', node.data)
                            else:
                                new = replace(node.data)
                            if new != node.data:
                                node.data = new
                                changed = True
                    # Names in sheet titles, attributes, or formula references need structural edits.
                    # Fail closed instead of releasing a workbook containing a registered name.
                    for el in doc.getElementsByTagName('*'):
                        values = [a.value for a in el.attributes.values()]
                        values += [c.data for c in el.childNodes if c.nodeType in (3, 4)]
                        if any(pattern.search(v) for v in values):
                            blocked.add(info.filename)
                    if changed:
                        raw = doc.toxml(encoding='utf-8')
                target.writestr(info, raw)
        if blocked:
            raise ValueError('シート名・数式参照・属性などの未対応箇所に登録名が残るため出力を停止しました。Excelで該当箇所を変更して再実行してください。対象部品: ' + ', '.join(sorted(blocked)[:5]))
    return output.getvalue(), counts, warnings

def process(payload):
    mapping = mapping_rules(payload.get('rows', []))
    files = payload.get('files', [])
    if not files or len(files) > 100:
        raise ValueError('Excelファイルを1〜100件選択してください。')
    bundle = io.BytesIO()
    reports = []
    with zipfile.ZipFile(bundle, 'w', zipfile.ZIP_DEFLATED) as archive:
        for index, item in enumerate(files, 1):
            report = {'file': str(item.get('name', '')), 'status': 'error'}
            try:
                if not report['file'].lower().endswith('.xlsx'):
                    raise ValueError('対応形式は.xlsxです。.xls・.xlsmはExcelで.xlsxとして保存してください。')
                data = base64.b64decode(item['data'], validate=True)
                result, counts, warnings = mask_xlsx(data, mapping)
                # Neutral output names avoid leaking personal names in paths or filenames.
                filename = f'masked_{index:03d}.xlsx'
                archive.writestr(filename, result)
                report.update(status='ok', output=filename, count=sum(counts.values()), counts=list(counts.values()), warnings=warnings)
            except Exception as exc:
                report['error'] = str(exc) if isinstance(exc, ValueError) else 'ファイルを読み込めません。破損・暗号化・ファイル形式を確認してください。'
            reports.append(report)
    return {'reports': reports, 'zip': base64.b64encode(bundle.getvalue()).decode() if any(r['status'] == 'ok' for r in reports) else None}

class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass
    def send(self, status, data, kind='application/json; charset=utf-8'):
        self.send_response(status)
        self.send_header('Content-Type', kind)
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)
    def do_GET(self):
        if self.path != '/':
            return self.send(404, b'{}')
        html = (ROOT / 'index.html').read_text(encoding='utf-8').replace('__TOKEN__', TOKEN)
        self.send(200, html.encode(), 'text/html; charset=utf-8')
    def do_POST(self):
        if self.path != '/api/mask' or self.headers.get('X-Mask-Token') != TOKEN:
            return self.send(403, b'{}')
        try:
            size = int(self.headers.get('Content-Length', '0'))
            if not 0 < size <= MAX_REQUEST:
                raise ValueError('一度に処理するファイルの合計は約100MB以下にしてください。')
            result = process(json.loads(self.rfile.read(size)))
            self.send(200, json.dumps(result, ensure_ascii=False).encode())
        except Exception as exc:
            self.send(400, json.dumps({'error': str(exc) if isinstance(exc, ValueError) else '処理に失敗しました。入力を確認してください。'}, ensure_ascii=False).encode())

if __name__ == '__main__':
    server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
    url = f'http://127.0.0.1:{server.server_port}/'
    print('Excel Masker: ' + url + '\n終了するにはこのウィンドウを閉じてください。', flush=True)
    threading.Timer(0.5, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        server.server_close()
