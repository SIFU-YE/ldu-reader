#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LDU-TGP 阅读难度评估 · 本地无限制版 Web 服务

- 基于 python-readability-cn 复刻的本地分析核心（core.py），无 1 万字限制
- 完全本地运行，不依赖任何远程服务
- 分级字/词/语法点明细全量展示
- 对外汉语使用 LTP 本地分析，母语汉语使用本地可读性公式

用法:
  python server.py                 # 默认 http://127.0.0.1:8001
  python server.py --port 9000
  curl -X POST http://127.0.0.1:8001/analyze -H "Content-Type: application/json" \
       -d '{"text":"...","country":"foreign"}'
"""

import argparse
import json
import os
import sys
import tempfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

import core
import report

sys.stdout.reconfigure(encoding="utf-8")

PAGE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>阅读难度评估 · 本地无限制版</title>
<style>
:root{--bg:#16161a;--panel:#1e1e24;--panel2:#26262e;--line:#3a3a44;--txt:#e8e8ea;--sub:#9a9aa4;--accent:#FAC775;--blue:#85B7EB;--green:#5DCAA5;--red:#F09595}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--txt);font-family:"Segoe UI","Microsoft YaHei",sans-serif;line-height:1.6;padding:24px}
.wrap{max-width:960px;margin:0 auto}
h1{font-size:20px;font-weight:500;margin-bottom:4px}
.sub{color:var(--sub);font-size:13px;margin-bottom:20px}
.panel{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:20px;margin-bottom:16px}
label{font-size:13px;color:var(--sub);margin-right:16px;cursor:pointer}
textarea{width:100%;min-height:200px;background:var(--panel2);color:var(--txt);border:1px solid var(--line);border-radius:8px;padding:12px;font-size:14px;resize:vertical;margin:12px 0;font-family:inherit}
button{background:var(--accent);color:#412402;border:none;border-radius:8px;padding:10px 28px;font-size:14px;font-weight:500;cursor:pointer}
button:hover{opacity:.9}
button:disabled{opacity:.5;cursor:wait}
.kpi{display:flex;gap:16px;flex-wrap:wrap;margin:12px 0}
.kpi .box{flex:1;min-width:150px;background:var(--panel2);border:1px solid var(--line);border-radius:10px;padding:14px}
.kpi .box .k{font-size:12px;color:var(--sub)}
.kpi .box .v{font-size:22px;font-weight:500;margin-top:2px}
.kpi .box .v.red{color:var(--red)}.kpi .box .v.blue{color:var(--blue)}.kpi .box .v.green{color:var(--green)}
table{width:100%;border-collapse:collapse;font-size:13px}
th,td{border:1px solid var(--line);padding:6px 10px;text-align:left;vertical-align:top}
th{background:var(--panel2);color:var(--sub);font-weight:500;white-space:nowrap}
td.tags{color:var(--sub);word-break:break-all}
h2{font-size:15px;font-weight:500;margin:14px 0 8px}
h2 span{color:var(--sub);font-size:12px;font-weight:400}
.err{color:var(--red);font-size:13px;margin-top:8px}
.note{color:var(--sub);font-size:12px}
#result{display:none}
</style>
</head>
<body>
<div class="wrap">
<h1>阅读难度评估 · 本地无限制版</h1>
<div class="sub">基于程勇 GF0025 公式（对外）与母语可读性公式（本地复刻）· 无 1 万字限制 · 分级字/词/语法点全量明细</div>
<div class="panel">
  <label><input type="radio" name="country" value="foreign" checked> 对外汉语</label>
  <label><input type="radio" name="country" value="native"> 母语汉语</label>
  <textarea id="text" placeholder="在此粘贴文章内容（无字数限制，支持任意长度）…"></textarea>
  <button id="run">开始分析</button>
  <button id="export" style="background:var(--blue);color:#042C53;margin-left:10px" disabled>导出 Word 报告</button>
  <div class="err" id="err"></div>
</div>
<div class="panel" id="result">
  <h2>分析结果</h2>
  <div class="kpi" id="kpi"></div>
  <div class="note" id="note"></div>
  <h2>分级字明细</h2><table id="t_char"></table>
  <h2>分级词明细</h2><table id="t_word"></table>
  <h2>词性兜底词明细</h2><table id="t_fallback"></table>
  <h2>超纲词与专名</h2><table id="t_beyond"></table>
  <h2>分级语法点明细</h2><table id="t_gram"></table>
  <h2>统计信息</h2><table id="t_stats"></table>
</div>
</div>
<script>
const $=id=>document.getElementById(id);
$('run').onclick=async()=>{
  const text=$('text').value.trim();
  const country=document.querySelector('input[name="country"]:checked').value;
  $('err').textContent='';
  if(!text){$('err').textContent='请输入文章内容';return}
  $('run').disabled=true;$('run').textContent='分析中…';
  try{
    const r=await (await fetch('/analyze',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({text,country})})).json();
    if(r.error) throw new Error(r.error);
    render(r);
  }catch(e){$('err').textContent='错误：'+e.message}
  $('run').disabled=false;$('run').textContent='开始分析';
};
function kpi(r){
  const c=[];const push=(k,v,cls)=>c.push(`<div class="box"><div class="k">${k}</div><div class="v ${cls}">${v}</div></div>`);
  if(r.country==='foreign'){
    push('难度分值',r.score,'red');push('对应级别',r.level,'blue');
    push('字数',r.stats.total_chars,'');push('句子数',r.stats.sentences,'');
  }else{
    push('可读性指数',r.score??'—','red');push('字数',r.stats.total_chars,'');push('句子数',r.stats.sentences,'');
  }
  $('kpi').innerHTML=c.join('');
}
const LV=['一级','二级','三级','四级','五级','六级','七-九级'];
function tags(arr){return arr&&arr.length?arr.join('、'):'<span class="tags">—</span>'}
function render(r){
  $('result').style.display='block';
  lastResult=r;
  $('export').disabled=false;
  kpi(r);
  $('note').textContent=r.note||'分级体系：1-1.4 一级 / 1.4-1.75 二级 / 1.75-2 三级 / 2-2.3 四级 / 2.3-2.55 五级 / 2.55-2.7 六级 / 2.7+ 七-九级';
  const rows=(arr)=>arr.map((v,i)=>`<tr><th>${LV[i]}</th><td class="tags">${tags(v)}</td></tr>`).join('');
  $('t_char').innerHTML='<tr><th style="width:90px">级别</th><th>汉字</th></tr>'+rows(r.char_detail);
  $('t_word').innerHTML='<tr><th style="width:90px">级别</th><th>词语</th></tr>'+rows(r.word_detail);
  $('t_fallback').innerHTML='<tr><th style="width:90px">级别</th><th>词语（词性不匹配·按词形定级）</th></tr>'+rows(r.word_fallback_detail||[]);
  const beyondRows=(key,label)=>{const v=r[key];return `<tr><th style="width:90px">${label}</th><td class="tags">${tags(Array.isArray(v)?v:Object.keys(v||{}))}</td></tr>`};
  $('t_beyond').innerHTML='<tr><th style="width:90px">类型</th><th>词语</th></tr>'
    +beyondRows('beyond_words_clean','普通超纲词')+beyondRows('beyond_proper_nouns','专名/缩略')
    +beyondRows('beyond_number_like','数字组合')+beyondRows('beyond_phrases','短语/分词结果')
    +beyondRows('beyond_fragments','单字残片');
  $('t_gram').innerHTML='<tr><th style="width:90px">级别</th><th>语法点</th></tr>'+rows(r.grammar_detail);
  const s=r.stats;
  $('t_stats').innerHTML=`<tr><th style="width:90px">总字数</th><td>${s.total_chars}</td><th style="width:90px">字种数</th><td>${s.unique_chars}</td></tr>
  <tr><th>句子数</th><td>${s.sentences}</td><th>平均句长</th><td>${s.avg_sent_len}</td></tr>
  <tr><th>分词Token数（含标点）</th><td>${s.total_words??'—'}</td><th>纯汉字词数</th><td>${s.total_hanzi_words??'—'}</td></tr>
  <tr><th>词种数</th><td>${s.unique_words??'—'}</td><th>模式</th><td>${r.country==='foreign'?'对外汉语':'母语汉语'}</td></tr>`;
}
let lastResult=null;
$('export').onclick=async()=>{
  if(!lastResult)return;
  $('export').disabled=true;$('export').textContent='生成中…';
  try{
    const resp=await fetch('/report',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({text:$('text').value.trim(),country:document.querySelector('input[name="country"]:checked').value})});
    if(!resp.ok)throw new Error(await resp.text());
    const blob=await resp.blob();
    const a=document.createElement('a');
    a.href=URL.createObjectURL(blob);a.download='阅读难度分析报告.docx';
    a.click();URL.revokeObjectURL(a.href);
  }catch(e){$('err').textContent='导出失败：'+e.message}
  $('export').disabled=false;$('export').textContent='导出 Word 报告';
};
</script>
</body>
</html>"""


def _analyze_local(text, country="foreign"):
    """完全本地分析，并补齐页面/报告所需的通用字段。"""
    if country == "foreign":
        return core.analyze_foreign_ltp(text)

    return core.analyze(text, "native")


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        sys.stderr.write("[%s] %s\n" % (self.log_date_time_string(), fmt % args))

    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        data = body.encode("utf-8") if isinstance(body, str) else body
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        if urlparse(self.path).path in ("/", "/index"):
            self._send(200, PAGE, "text/html; charset=utf-8")
        else:
            self._send(404, json.dumps({"error": "not found"}))

    def do_POST(self):
        path = urlparse(self.path).path
        if path not in ("/analyze", "/report"):
            self._send(404, json.dumps({"error": "not found"}))
            return
        try:
            length = int(self.headers.get("Content-Length", 0))
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            text = payload.get("text", "")
            country = payload.get("country", "foreign")
            if not text.strip():
                raise ValueError("输入为空")
            if path == "/analyze":
                result = _analyze_local(text, country)
                self._send(200, json.dumps(result, ensure_ascii=False))
            else:
                # 导出 Word 报告
                import io
                buf = io.BytesIO()
                out_path = os.path.join(tempfile.gettempdir(), "readability_report.docx")
                result = _analyze_local(text, country)
                report.make_report(result, out_path)
                with open(out_path, "rb") as f:
                    data = f.read()
                self.send_response(200)
                self.send_header("Content-Type",
                                 "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                self.send_header("Content-Disposition",
                                 'attachment; filename="readability_report.docx"')
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
        except ValueError as exc:
            self._send(400, json.dumps({"error": str(exc)}))
        except Exception as exc:  # noqa: BLE001
            self._send(500, json.dumps({"error": f"{type(exc).__name__}: {exc}"}))


def main():
    ap = argparse.ArgumentParser(description="LDU-TGP 阅读难度评估 · 本地无限制版")
    ap.add_argument("--port", type=int, default=8001, help="监听端口，默认 8001")
    ap.add_argument("--host", default="127.0.0.1", help="监听地址，默认 127.0.0.1")
    ap.add_argument("--no-browser", action="store_true", help="启动后不自动打开浏览器")
    args = ap.parse_args()
    try:
        import ltp  # noqa: F401  对外汉语 LTP 本地分析
    except ImportError:
        print("缺少依赖：请先执行 pip install ltp torch", file=sys.stderr)
        sys.exit(1)
    import threading
    import webbrowser
    url = f"http://{args.host}:{args.port}"
    print(f"阅读难度评估服务已启动: {url}  (关闭本窗口或 Ctrl+C 停止)")
    if not args.no_browser:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    ThreadingHTTPServer((args.host, args.port), Handler).serve_forever()


if __name__ == "__main__":
    main()
