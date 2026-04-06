#!/usr/bin/env python3
"""
优质中小企业自检报告 PDF 生成器（自举式 Playwright 版）
无需手动配置环境，内置自愈下载能力，跨平台部署。
"""

import sys
import subprocess
import os
import re

def ensure_dependencies():
    """环境自愈检查：检测到缺失将自动启动傻瓜式后台安转"""
    need_restart = False
    try:
        import markdown
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("💡 [环境自愈] 首次执行，正在为您静默下载轻量级 PDF 渲染引擎 (Playwright & Markdown)...")
        print("⏳ 请耐心等待 1-2 分钟，这仅会在首次运行时发生。")
        try:
            # 静默安装 markdown 和 playwright
            subprocess.check_call([sys.executable, "-m", "pip", "install", "markdown", "playwright"], stdout=subprocess.DEVNULL)
            print("🚀 前沿渲染核心加载完毕！正在后台为您静默组装专用的零负担浏览器渲染室 (Chromium 内核)...")
            # 自动化下载无头浏览器二进制包
            subprocess.check_call([sys.executable, "-m", "playwright", "install", "chromium"], stdout=subprocess.DEVNULL)
            print("✨ 跨平台自给足核准完毕！开始转印高阶咨询风 PDF...\n")
            need_restart = True
        except subprocess.CalledProcessError as e:
            print(f"❌ 内核框架部署受到系统阻力：{e}")
            sys.exit(1)
            
    if need_restart:
        # 依赖装载完毕后，为了保证路径和动态变量刷新，系统重发当前脚本。
        os.execv(sys.executable, [sys.executable] + sys.argv)

# 在 import 高阶库前，必须强制走完这道安检门
ensure_dependencies()

import markdown
from playwright.sync_api import sync_playwright
from pathlib import Path

def get_base_css():
    """极简现代的四大高阶咨询风 CSS"""
    return '''
        body {
            font-family: "PingFang SC", "Microsoft YaHei", "Source Han Sans CN", sans-serif;
            font-size: 11pt;
            line-height: 1.7;
            color: #1e293b;
            text-align: justify;
            margin: 0;
            padding: 0;
        }
        
        /* 封面级大标题 */
        h1 {
            font-size: 24pt;
            font-family: "Songti SC", "SimSun", serif;
            color: #0f172a;
            text-align: center;
            font-weight: 600;
            padding-bottom: 0.8cm;
            margin-top: 0;
            margin-bottom: 1.2cm;
            border-bottom: 1px solid #cbd5e1;
            page-break-after: avoid;
            letter-spacing: 2px;
        }
        
        /* 章节标题 */
        h2 {
            font-size: 15pt;
            color: #ffffff;
            background-color: #0f172a;
            padding: 0.25cm 0.5cm;
            margin-top: 1.2cm;
            margin-bottom: 0.8cm;
            page-break-after: avoid;
            border-radius: 2px;
            letter-spacing: 1px;
        }
        
        /* 次级标题 */
        h3 {
            font-size: 13pt;
            color: #334155;
            margin-top: 1cm;
            margin-bottom: 0.5cm;
            border-left: 3px solid #3b82f6;
            padding-left: 0.3cm;
            page-break-after: avoid;
        }
        
        /* 表格高级咨询风 (水平线为主) */
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 0.8cm 0;
            font-size: 9.5pt;
            page-break-inside: avoid;
        }
        
        th {
            background-color: #f8fafc;
            color: #0f172a;
            padding: 0.4cm 0.3cm;
            text-align: left;
            font-weight: bold;
            border-top: 2px solid #0f172a;
            border-bottom: 1px solid #94a3b8;
        }
        
        td {
            padding: 0.35cm 0.3cm;
            border-bottom: 1px solid #e2e8f0;
            vertical-align: top;
            color: #334155;
        }
        
        tr:last-child td {
            border-bottom: 2px solid #cbd5e1;
        }
        
        /* 行内强调 */
        strong {
            color: #0f172a;
            font-weight: 600;
        }
        
        em {
            color: #0ea5e9;
            font-style: italic;
        }
        
        /* 列表排版 */
        ul, ol {
            margin: 0.4cm 0;
            padding-left: 0.8cm;
            color: #334155;
        }
        
        li {
            margin: 0.2cm 0;
        }
        
        /* 警告与提示框 (Callouts) */
        .callout {
            padding: 0.5cm 0.6cm;
            margin: 0.6cm 0;
            border-radius: 4px;
            page-break-inside: avoid;
            font-size: 10pt;
        }
        
        .callout-warning, .callout-caution {
            background-color: #fffbeb;
            border-left: 4px solid #f59e0b;
            color: #b45309;
        }
        
        .callout-note, .callout-info {
            background-color: #f0fdf4;
            border-left: 4px solid #10b981;
            color: #065f46;
        }
        
        .callout-important {
            background-color: #eff6ff;
            border-left: 4px solid #3b82f6;
            color: #1e40af;
        }
        
        /* 默认 Blockquote 样式 */
        blockquote {
            border-left: 3px solid #cbd5e1;
            padding-left: 0.5cm;
            margin: 0.6cm 0;
            color: #64748b;
            background-color: #f8fafc;
            padding: 0.4cm 0.5cm;
            font-style: italic;
        }
        
        /* 代码块 */
        code {
            background-color: #f1f5f9;
            padding: 0.05cm 0.15cm;
            border-radius: 2px;
            font-family: Consolas, monospace;
            font-size: 9pt;
            color: #e11d48;
        }
        
        pre {
            background-color: #0f172a;
            color: #f8fafc;
            padding: 0.5cm;
            border-radius: 4px;
            font-size: 8.5pt;
            page-break-inside: avoid;
            font-family: Consolas, monospace;
        }
        
        pre code {
            background-color: transparent;
            color: inherit;
            padding: 0;
        }
        
        hr {
            border: none;
            border-top: 1px solid #cbd5e1;
            margin: 1cm 0;
        }
    '''

def markdown_to_html(md_content):
    """预处理 GitHub Alerts 并转换为 HTML"""
    def callout_repl(match):
        alert_type = match.group(1).lower()
        content = match.group(2).strip()
        html_content = markdown.markdown(content, extensions=['tables'])
        return f'<div class="callout callout-{alert_type}"><strong>{alert_type.upper()}</strong><br/>{html_content}</div>\n\n'

    md_content = re.sub(
        r'> \[!(WARNING|CAUTION|NOTE|IMPORTANT|TIP|INFO)\]\n((?:>.*\n)*)', 
        lambda m: callout_repl(re.match(r'> \[!(\w+)\]\n((?:>.*\n?)*)', m.group(0))), 
        md_content
    )
    md_content = md_content.replace('> ', '')
    
    html = markdown.markdown(
        md_content,
        extensions=['tables', 'fenced_code', 'codehilite', 'toc', 'nl2br']
    )
    
    full_html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <style>{get_base_css()}</style>
</head>
<body>
{html}
</body>
</html>'''
    return full_html

def generate_pdf(input_md, output_pdf):
    input_path = Path(input_md)
    if not input_path.exists():
        print(f"错误：文件不见了？无法读取 - {input_md}")
        return False
    
    md_content = input_path.read_text(encoding='utf-8')
    html_content = markdown_to_html(md_content)
    output_path = Path(output_pdf)
    
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()
            page.set_content(html_content)
            
            # 使用 Playwright 的原生接口来投射精美的固定页眉和记分牌页脚
            header_template = '<div style="font-size:8pt; color:#94a3b8; margin-left:2.2cm; font-family:\'PingFang SC\', sans-serif;">优质中小企业多维定性与穿透诊断底稿</div>'
            footer_template = '<div style="font-size:9pt; color:#64748b; text-align:center; width:100%; font-family:Helvetica, Arial, sans-serif;"><span class="pageNumber"></span> / <span class="totalPages"></span></div>'
            
            page.pdf(
                path=str(output_path),
                format="A4",
                margin={"top": "2.5cm", "bottom": "2.5cm", "left": "2.2cm", "right": "2.2cm"},
                display_header_footer=True,
                header_template=header_template,
                footer_template=footer_template,
                print_background=True
            )
            browser.close()
            
        print(f"✅ PDF 诊断报告已生成：{output_pdf}")
        return True
    except Exception as e:
        # 捕捉意外情况
        if "Executable doesn't exist" in str(e):
            print("⚠️ 奇怪，无头浏览器突然失联。建议紧急运行：python -m playwright install chromium")
        else:
            print(f"❌ 渲染底座塌陷！：{e}")
        return False

def main():
    if len(sys.argv) < 3:
        print("用法：python generate_diagnosis_pdf.py <体检报告.md> <输出.pdf>")
        sys.exit(1)
    
    input_md = sys.argv[1]
    output_pdf = sys.argv[2]
    
    Path(output_pdf).parent.mkdir(parents=True, exist_ok=True)
    generate_pdf(input_md, output_pdf)

if __name__ == '__main__':
    main()
