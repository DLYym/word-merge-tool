# -*- coding: utf-8 -*-
"""
Word 批量替换合并工具（Windows）
=================================
功能：
    1. 选择一个 Word 模板文档（.docx）
    2. 指定“待替换的词”（例如：朋友）
    3. 输入“替换后的词”列表（每行一个，例如：家人 / 老师 / 同事 / 同学）
    4. 点击“开始合并”，程序会为每个替换词“另存一份模板副本”，对每份副本做
       全文替换，最后把所有结果按顺序合并成一个 Word 文档（每份之间换页）。

实现思路（完全模拟手动操作）：
    - 用 Documents.Add(Template=模板) 为每个替换词新建一份副本，完整继承
      模板的页面设置、图片/浮动对象位置、页眉页脚（避免复制粘贴导致对象错位）。
    - 对每份副本做“全文替换”（整份正文查找替换），避免漏替换。
    - 从第 2 份起，给其第一段设置“段前分页”，合并后每份自动另起一页，
      且不会像插入分页符那样在下一页顶部残留空白行。

依赖（Windows）：
    pip install pywin32

打包成 exe（Windows）：
    pip install pyinstaller
    pyinstaller --noconsole --onefile --name "WordMergeTool" word_merge_tool.py
"""

import os
import queue
import threading

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# ---------------------------------------------------------------------------
# win32com 依赖检测（该库仅 Windows 可用）
# ---------------------------------------------------------------------------
try:
    import pythoncom
    import win32com.client as win32
    HAS_WIN32COM = True
except ImportError:
    pythoncom = None
    win32 = None
    HAS_WIN32COM = False

# 说明：这里用 win32.Dispatch（后期绑定）操作 Word，访问 wdConst.xxx 常量
# 在后期绑定下不稳定（会报 AttributeError: <unknown>.xxx），所以全部改用
# 字面整数值，最稳妥。
WD_FORMAT_DOCX = 16    # wdFormatXMLDocument，对应 .docx
WD_REPLACE_ALL = 2     # wdReplaceAll
WD_ALERTS_NONE = 0     # wdAlertsNone，关闭 Word 弹窗
WD_FIND_CONTINUE = 1   # wdFindContinue，整篇循环查找
WD_COLLAPSE_END = 0    # wdCollapseEnd，折叠到末尾


# ---------------------------------------------------------------------------
# 核心处理逻辑（在后台线程中运行，避免卡死界面）
# ---------------------------------------------------------------------------
def merge_documents(template_path, output_path, replace_list, target_word,
                    show_word=False, log=None):
    """完全模拟手动操作：为每个替换词另存一份模板副本并全文替换，再合并。"""
    if not HAS_WIN32COM:
        raise RuntimeError("当前环境缺少 pywin32，请先在 Windows 上执行：pip install pywin32")

    if not os.path.exists(template_path):
        raise FileNotFoundError("模板文件不存在：" + template_path)

    if not replace_list:
        raise ValueError("替换词列表不能为空")

    if os.path.abspath(output_path) == os.path.abspath(template_path):
        raise ValueError("输出文件不能与模板文件相同，请换一个输出路径")

    # COM 在线程中使用必须先初始化；由于要用到剪贴板（Copy/Paste），
    # 线程必须以 STA（单线程单元）模式初始化，否则剪贴板操作会失败。
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)

    word = None
    src_doc = None
    dst_doc = None
    temp_docs = []
    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = show_word
        word.DisplayAlerts = WD_ALERTS_NONE  # 关闭弹窗，避免卡住后台线程

        if log:
            log("正在打开模板：" + os.path.basename(template_path))

        # 打开模板（只读）。每份副本下面用 Documents.Add(Template=...) 基于
        # 模板文件新建，从而完整继承页面设置、图片/浮动对象位置、页眉页脚。
        src_doc = word.Documents.Open(template_path, ReadOnly=True)

        total = len(replace_list)

        # 第一步：为每个替换词生成一份独立的、替换好的文档。
        for i, new_word in enumerate(replace_list, start=1):
            if log:
                log("[%d/%d] 生成副本并全文替换：%s" % (i, total, new_word))

            # 以模板为基础新建一份副本（等价“另存为一份”），这是解决
            # “图片/对象位置错乱”的关键：页面设置与模板完全一致。
            temp_doc = word.Documents.Add(Template=template_path)

            # 全文替换（模拟 Ctrl+H 全部替换）：范围是整份文档正文。
            # 用整份 Content.Find 替换，而不是只在“刚粘贴的一段”里替换，
            # 能避免漏替换（解决“部分内容不对”）。
            find = temp_doc.Content.Find
            find.ClearFormatting()
            # FindText / ReplaceWith 必须显式传给 Execute，否则会被默认
            # 空参数覆盖，导致替换不生效。
            find.Execute(
                FindText=target_word,
                MatchCase=False,
                MatchWholeWord=False,
                MatchWildcards=False,
                MatchSoundsLike=False,
                MatchAllWordForms=False,
                Forward=True,
                Wrap=WD_FIND_CONTINUE,  # 整篇循环查找替换
                Format=False,
                ReplaceWith=new_word,
                Replace=WD_REPLACE_ALL,
            )

            # 从第二份开始，给第一段设置“段前分页”，合并后每份自动另起一页，
            # 且不会像插入分页符那样在下一页顶部残留空白行。
            # 注意：不要用 .Paragraphs(1)（win32com 下返回对象无
            # ParagraphFormat 属性）；直接用 collapsed Range，作用于所在整段。
            # if i > 1:
            #     # temp_doc.Range(0, 0).ParagraphFormat.PageBreakBefore = True
            #     temp_doc.Paragraphs(1).Format.PageBreakBefore = True

            temp_docs.append(temp_doc)

        # 第二步：把所有独立文档合并成一个。
        # 第一份直接作为最终文档基础（已继承模板页面设置），其余依次追加。
        dst_doc = temp_docs[0]

        for i in range(1, total):
            if log:
                log("[%d/%d] 合并第 %d 份" % (i + 1, total, i + 1))

            # ---------- 在目标文档末尾插入分页符 ----------
            # 将光标定位到文档末尾（最后一个段落标记之前）
            end_range = dst_doc.Content
            end_range.Collapse(WD_COLLAPSE_END)   # 折叠到末尾
            # end_range.InsertBreak(7)              # 7 = wdPageBreak，插入分页符
            end_range.Text = "\f"         #插入分页符字符
            # ------------------------------------------------
            
            # 复制：激活源文档并“全选复制”。
            # 关键：不能用 Content.Copy()（Range.Copy 只复制文本流，
            # 会丢失锚定在段落上的“浮动图片”），而 Selection.WholeStory()
            # 模拟手动 Ctrl+A 全选，能连同浮动图片一起复制——这正是
            # 表格里“联系人类型”列那些方框/选项图片（浮动图片）能
            # 正常复制过去的原因。
            temp_docs[i].Activate()
            word.Selection.WholeStory()
            word.Selection.Copy()

            # 粘贴：目标文档光标移到末尾，粘贴整份内容
            dst_doc.Activate()
            rng = dst_doc.Content
            rng.Collapse(WD_COLLAPSE_END)
            rng.Select()
            word.Selection.Paste()

        if log:
            log("正在保存结果文件 ...")
        dst_doc.SaveAs2(output_path, FileFormat=WD_FORMAT_DOCX)

        # 清除合并过程中写入剪贴板的内容
        try:
            word.CutCopyMode = False
        except Exception:
            pass
        try:
            import win32clipboard
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.CloseClipboard()
        except Exception:
            pass

        if log:
            log("完成！已生成：" + output_path)

    except Exception as e:
        raise RuntimeError(str(e)) from e
    finally:
        # 关闭所有文档并退出 Word，确保进程被释放
        for d in temp_docs:
            try:
                if d is not None:
                    d.Close(False)
            except Exception:
                pass
        try:
            if src_doc is not None:
                src_doc.Close(False)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


# ---------------------------------------------------------------------------
# 图形界面
# ---------------------------------------------------------------------------
class WordMergeApp:
    def __init__(self, root):
        self.root = root
        self.msg_queue = queue.Queue()

        root.title("Word 批量替换合并工具")
        root.geometry("680x560")
        root.minsize(640, 520)

        self.template_var = tk.StringVar()
        self.target_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.show_word_var = tk.BooleanVar(value=False)

        self._build_ui()

        # 非 Windows 环境给出提示（但界面仍可打开查看）
        if not HAS_WIN32COM:
            self.log("【提示】当前环境未检测到 pywin32，本工具需在 Windows 上运行。")
            self.log("        Windows 上请先执行：pip install pywin32")

        self.root.after(100, self._poll_queue)

    # ------------------------------- UI 布局 -------------------------------
    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill="both", expand=True)
        main.columnconfigure(1, weight=1)

        # 1. 模板文档
        ttk.Label(main, text="模板文档：").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(main, textvariable=self.template_var).grid(
            row=0, column=1, sticky="ew", **pad)
        ttk.Button(main, text="浏览...", command=self._pick_template).grid(
            row=0, column=2, **pad)

        # 2. 待替换的词
        ttk.Label(main, text="待替换的词：").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(main, textvariable=self.target_var).grid(
            row=1, column=1, sticky="ew", **pad)
        ttk.Label(main, text="（例如：朋友）").grid(row=1, column=2, sticky="w", **pad)

        # 3. 替换后的词（多行，每行一个）
        ttk.Label(main, text="替换后的词：").grid(
            row=2, column=0, sticky="nw", **pad)
        self.replace_text = tk.Text(main, height=10, width=40)
        self.replace_text.grid(row=2, column=1, columnspan=2, sticky="nsew", **pad)
        main.rowconfigure(2, weight=1)

        # 3.1 导入替换词
        ttk.Button(main, text="从文件导入替换词...",
                   command=self._import_words).grid(
            row=3, column=1, sticky="w", padx=10, pady=(0, 6))
        ttk.Label(main, text="每行一个词；支持 .txt / .csv（取第一列）").grid(
            row=3, column=2, sticky="w", pady=(0, 6))

        # 4. 输出文件
        ttk.Label(main, text="输出文件：").grid(row=4, column=0, sticky="w", **pad)
        ttk.Entry(main, textvariable=self.output_var).grid(
            row=4, column=1, sticky="ew", **pad)
        ttk.Button(main, text="浏览...", command=self._pick_output).grid(
            row=4, column=2, **pad)

        # 5. 选项 + 开始按钮
        opt_row = ttk.Frame(main)
        opt_row.grid(row=5, column=0, columnspan=3, sticky="w", **pad)
        ttk.Checkbutton(opt_row, text="显示 Word 界面（调试用）",
                        variable=self.show_word_var).pack(side="left")
        self.start_btn = ttk.Button(opt_row, text="开始合并",
                                    command=self._on_start)
        self.start_btn.pack(side="right")

        # 6. 日志区
        ttk.Label(main, text="运行日志：").grid(row=6, column=0, sticky="nw", **pad)
        self.log_text = scrolledtext.ScrolledText(main, height=8, state="disabled")
        self.log_text.grid(row=6, column=1, columnspan=2, sticky="nsew", **pad)
        main.rowconfigure(6, weight=1)

    # ------------------------------- 交互逻辑 -------------------------------
    def _pick_template(self):
        path = filedialog.askopenfilename(
            title="选择模板 Word 文档",
            filetypes=[("Word 文档", "*.docx *.doc"), ("所有文件", "*.*")])
        if path:
            self.template_var.set(path)
            # 默认输出到模板同目录
            if not self.output_var.get():
                out = os.path.splitext(path)[0] + "_合并结果.docx"
                self.output_var.set(out)

    def _pick_output(self):
        path = filedialog.asksaveasfilename(
            title="保存合并结果",
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")])
        if path:
            self.output_var.set(path)

    def _import_words(self):
        path = filedialog.askopenfilename(
            title="选择替换词文件",
            filetypes=[("文本文件", "*.txt"), ("CSV 文件", "*.csv"),
                       ("所有文件", "*.*")])
        if not path:
            return
        try:
            words = []
            with open(path, "r", encoding="utf-8-sig") as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    # 若为 csv，取第一列
                    if "," in line:
                        line = line.split(",")[0].strip()
                    if line:
                        words.append(line)
            self.replace_text.delete("1.0", "end")
            self.replace_text.insert("1.0", "\n".join(words))
            self.log("已从文件导入 %d 个替换词。" % len(words))
        except Exception as e:
            messagebox.showerror("导入失败", str(e))

    def _on_start(self):
        template_path = self.template_var.get().strip()
        target_word = self.target_var.get().strip()
        output_path = self.output_var.get().strip()

        # 校验输入
        if not template_path:
            messagebox.showwarning("提示", "请先选择模板 Word 文档。")
            return
        if not target_word:
            messagebox.showwarning("提示", "请输入“待替换的词”。")
            return
        replace_list = [w.strip() for w in self.replace_text.get("1.0", "end").splitlines()
                        if w.strip()]
        if not replace_list:
            messagebox.showwarning("提示", "请至少输入一个“替换后的词”（每行一个）。")
            return
        if not output_path:
            messagebox.showwarning("提示", "请指定输出文件路径。")
            return
        if not output_path.lower().endswith(".docx"):
            output_path += ".docx"
            self.output_var.set(output_path)

        # 禁用按钮，启动后台线程
        self.start_btn.config(state="disabled")
        self.log("=" * 50)
        self.log("开始合并，共 %d 个替换词。" % len(replace_list))
        threading.Thread(
            target=self._worker,
            args=(template_path, output_path, replace_list, target_word),
            daemon=True,
        ).start()

    def _worker(self, template_path, output_path, replace_list, target_word):
        try:
            merge_documents(
                template_path, output_path, replace_list, target_word,
                show_word=self.show_word_var.get(), log=self.log)
        except Exception as e:
            self.log("【出错】" + str(e))
            # 通过队列把弹窗请求传回主线程
            self.msg_queue.put(("error", "合并失败：" + str(e)))
        finally:
            self.msg_queue.put(("done", None))

    # ------------------------------- 日志/消息 -------------------------------
    def log(self, msg):
        self.msg_queue.put(("log", msg))

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.msg_queue.get_nowait()
                if kind == "log":
                    self.log_text.config(state="normal")
                    self.log_text.insert("end", payload + "\n")
                    self.log_text.see("end")
                    self.log_text.config(state="disabled")
                elif kind == "error":
                    messagebox.showerror("错误", payload)
                elif kind == "done":
                    self.start_btn.config(state="normal")
                    self.log("任务结束。")
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)


# ---------------------------------------------------------------------------
def main():
    root = tk.Tk()
    WordMergeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
