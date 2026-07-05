import os
import threading
import tkinter as tk
from tkinter import messagebox, scrolledtext, simpledialog
import re
import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

# ---------- User‑Agent 管理 ----------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
UA_FILE = os.path.join(SCRIPT_DIR, 'UA.txt')

DEFAULT_UA = (
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
    'AppleWebKit/537.36 (KHTML, like Gecko) '
    'Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0'
)

# 全局请求头，后续会在 start() 中根据有效 UA 更新
HEADERS = {
    'User-Agent': DEFAULT_UA,
    'Accept-Language': 'zh-CN,zh;q=0.9'
}

def is_valid_ua(ua_str):
    """User‑Agent 至少需要 20 个字符才视为合理"""
    return bool(ua_str) and len(ua_str.strip()) >= 20

def load_user_agent():
    """尝试从文件读取有效的 User‑Agent，失败返回 None"""
    try:
        if os.path.exists(UA_FILE):
            with open(UA_FILE, 'r', encoding='utf-8') as f:
                ua = f.read().strip()
                return ua if is_valid_ua(ua) else None
        return None
    except Exception:
        return None

def save_user_agent(ua):
    """将 User‑Agent 写入文件，返回是否成功"""
    try:
        with open(UA_FILE, 'w', encoding='utf-8') as f:
            f.write(ua.strip())
        return True
    except Exception:
        return False

def get_desktop_path():
    """返回当前用户桌面路径，若获取失败则返回脚本所在目录"""
    try:
        # Windows / macOS / Linux 通用写法
        desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
        if os.path.exists(desktop):
            return desktop
    except Exception:
        pass
    # 回退到脚本所在目录
    return SCRIPT_DIR

# ---------- 爬虫 ----------
def fetch_word_info(word):
    url = f'https://www.youdao.com/result?word={word}&lang=en'
    try:
        resp = requests.get(url, headers=HEADERS, timeout=10)
        resp.raise_for_status()
    except requests.RequestException as e:
        raise RuntimeError(f"网络请求失败: {e}")

    soup = BeautifulSoup(resp.text, 'html.parser')

    # 单词本体（只取直接文本）
    title_tag = soup.find('div', class_='title')
    if title_tag:
        direct_text = title_tag.find(string=True, recursive=False)  # 已修复弃用警告
        actual_word = direct_text.strip() if direct_text else word
    else:
        actual_word = word

    # 词性 + 释义
    pos_trans = []
    basic_ul = soup.select_one('.trans-container .basic')
    if basic_ul:
        for li in basic_ul.find_all('li', class_='word-exp'):
            pos_tag = li.find('span', class_='pos')
            trans_tag = li.find('span', class_='trans')
            if pos_tag and trans_tag:
                pos = pos_tag.get_text(strip=True)
                trans = trans_tag.get_text(strip=True)
                pos_trans.append((pos, trans))

    # 双语例句
    sentences = []
    sents_div = soup.find('div', class_='blng_sents_part')
    if sents_div:
        for li in sents_div.select('ul li.mcols-layout'):
            eng_tag = li.select_one('.sen-eng')
            chn_tag = li.select_one('.sen-ch')
            if eng_tag and chn_tag:
                eng = eng_tag.get_text(strip=True)
                chn = chn_tag.get_text(strip=True)
                sentences.append((eng, chn))

    return {
        'word': actual_word,
        'pos_trans': pos_trans,
        'sentences': sentences
    }

# ---------- GUI ----------
class WordApp:
    def __init__(self, root):
        self.root = root
        root.title("气死雷颜器")
        root.geometry("500x400")
        root.resizable(False, False)

        icon_path = os.path.join(SCRIPT_DIR, 'starbucks.ico')
        try:
            root.iconbitmap(icon_path)
        except Exception:
            pass   # 图标加载失败不影响程序运行

        tk.Label(root, text="输入单词（每行一个，或用逗号/空格分隔）：").pack(pady=(10, 0))
        self.word_text = scrolledtext.ScrolledText(root, height=10, width=60)
        self.word_text.pack(pady=5)

        limit_frame = tk.Frame(root)
        limit_frame.pack(pady=5)
        tk.Label(limit_frame, text="释义个数上限（0 为不限制）：").pack(side=tk.LEFT)
        self.limit_entry = tk.Entry(limit_frame, width=5)
        self.limit_entry.insert(0, "0")
        self.limit_entry.pack(side=tk.LEFT, padx=5)

        self.start_btn = tk.Button(root, text="开始生成 Word 文档", command=self.start)
        self.start_btn.pack(pady=10)

        self.status_label = tk.Label(root, text="就绪", fg="gray")
        self.status_label.pack()
        
        self.info_label = tk.Label(root, text="本程序由杨镇源开发，仅供完成雷颜布置的20篇20词积累使用。\n请勿向雷颜透露、将本程序另作他用等，最终解释权归作者本人所有。\n© anghenuan", fg="gray")
        self.info_label.pack()

    def parse_words(self):
        raw = self.word_text.get("1.0", tk.END)
        words = re.split(r'[,\s]+', raw)
        return [w.strip() for w in words if w.strip()]

    def start(self):
        # ----- 强制检查 User‑Agent -----
        ua = load_user_agent()
        while ua is None:
            # 弹出输入框，要求填写 UA
            user_input = simpledialog.askstring(
                "User‑Agent 配置",
                "未找到有效的 User‑Agent 文件。\n请输入完整的 User‑Agent 字符串（至少 20 个字符）：",
                parent=self.root
            )
            if user_input is None:  # 用户点了取消
                messagebox.showwarning("操作取消", "未配置 User‑Agent，无法继续运行。")
                return
            user_input = user_input.strip()
            if not is_valid_ua(user_input):
                messagebox.showwarning("格式错误", "User‑Agent 长度不足或为空，请重新输入。")
                continue
            if save_user_agent(user_input):
                ua = user_input
                break
            else:
                messagebox.showerror("保存失败", "无法写入 user_agent.txt，请检查文件权限。")
                return

        # 更新全局请求头
        global HEADERS
        HEADERS['User-Agent'] = ua

        # ----- 原有抓取流程 -----
        words = self.parse_words()
        if not words:
            messagebox.showwarning("提示", "请至少输入一个单词")
            return
        try:
            limit = int(self.limit_entry.get().strip())
        except ValueError:
            limit = 0

        self.start_btn.config(state=tk.DISABLED)
        self.status_label.config(text="正在抓取中...")
        threading.Thread(target=self.run, args=(words, limit), daemon=True).start()

    def run(self, words, limit):
        results = []
        total = len(words)
        for i, w in enumerate(words, 1):
            self.status_label.config(text=f"正在处理 ({i}/{total}): {w}")
            try:
                info = fetch_word_info(w)
                pos_trans = []
                for pos, trans in info['pos_trans']:
                    if limit > 0:
                        items = trans.split('；')
                        trans = '；'.join(items[:limit])
                    pos_trans.append((pos, trans))
                results.append({
                    'index': i,
                    'word': info['word'],
                    'pos_trans': pos_trans,
                    'sentences': info['sentences']
                })
            except Exception as e:
                results.append({
                    'index': i,
                    'word': w,
                    'pos_trans': [],
                    'sentences': [],
                    'error': str(e)
                })
        self.generate_word(results)
        self.status_label.config(text="生成完毕，文件已保存为“单词积累.docx”")
        self.start_btn.config(state=tk.NORMAL)
        messagebox.showinfo("完成", "Word 文档已生成在桌面：单词积累.docx")

    def generate_word(self, entries):
        doc = Document()
        style = doc.styles['Normal']
        font = style.font
        font.name = '等线'
        font.size = Pt(11)
        style.element.rPr.rFonts.set(qn('w:eastAsia'), '等线')

        for entry in entries:
            p = doc.add_paragraph()
            run_title = p.add_run(f"{entry['index']}. {entry['word']}   ")
            run_title.bold = True
            self.set_font(run_title)

            if 'error' in entry:
                run_err = p.add_run(f"（抓取失败: {entry['error']}）")
                self.set_font(run_err)
                continue

            for pos, trans in entry['pos_trans']:
                run_pos = p.add_run(f" {pos} ")
                run_pos.bold = True
                self.set_font(run_pos)
                run_trans = p.add_run(f"{trans}  ")
                self.set_font(run_trans)

            if entry['sentences']:
                for idx, (eng, chn) in enumerate(entry['sentences'], 1):
                    p_ex = doc.add_paragraph()
                    p_ex.paragraph_format.left_indent = Pt(20)
                    run_label = p_ex.add_run(f"例{idx}: ")
                    run_label.bold = True
                    self.set_font(run_label)
                    run_content = p_ex.add_run(f"{eng}  {chn}")
                    self.set_font(run_content)
            else:
                p_ex = doc.add_paragraph()
                p_ex.paragraph_format.left_indent = Pt(20)
                run_label = p_ex.add_run("例: ")
                run_label.bold = True
                self.set_font(run_label)
                run_none = p_ex.add_run("（无例句）")
                self.set_font(run_none)

        output_path = os.path.join(get_desktop_path(), "单词积累.docx")
        doc.save(output_path)

    def set_font(self, run):
        run.font.name = '等线'
        run.font.size = Pt(11)
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '等线')

if __name__ == "__main__":
    root = tk.Tk()
    app = WordApp(root)
    root.mainloop()
