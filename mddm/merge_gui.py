"""Desktop interface for the Word document merger."""

import os
from pathlib import Path
from queue import Empty, Queue
import subprocess
import sys
from threading import Thread
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


BACKGROUND = "#EDF2F8"
PANEL = "#FFFFFF"
BORDER = "#D8E2EF"
INK = "#182D47"
MUTED = "#526982"
BLUE = "#194EB8"
GREEN = "#14684F"
RED = "#AB3433"
CONSOLE = "#15263E"
DOCX_FILTER = [("Word 文档", "*.docx")]


def validate_paths(original, copies, output):
    if not all(value.strip() for value in (original, copies, output)):
        raise ValueError("请选择原稿、副本目录和合稿保存位置。")

    original, copies, output = (
        Path(value.strip()).expanduser().resolve() for value in (original, copies, output)
    )
    if original.suffix.lower() != ".docx" or not original.is_file():
        raise ValueError("请选择一个存在的 .docx 原稿。")
    if not copies.is_dir():
        raise ValueError("请选择包含部门副本的目录。")
    if output.suffix.lower() != ".docx" or output.is_dir():
        raise ValueError("合稿保存位置须为 .docx 文件。")

    copy_paths = [
        path for path in copies.iterdir()
        if path.is_file() and path.suffix.lower() == ".docx"
        and not path.name.startswith("~$") and not path.samefile(original)
    ]
    if not copy_paths:
        raise ValueError("副本目录中没有可合并的 .docx 文件。")
    if output.exists() and any(output.samefile(path) for path in [original, *copy_paths]):
        raise ValueError("请为合稿选择单独的保存位置，以免覆盖原稿或副本。")
    return original, copies, output


def result_paths(output):
    return (
        output,
        output.with_name(f"{output.stem}_conflicts.txt"),
        output.with_name(f"{output.stem}_merged_changes_log.txt"),
    )


def run_merge(paths, events):
    """Stream the existing CLI output to the Tk event queue."""
    original, copies, output = paths
    python = Path(sys.executable)
    if python.name.lower() == "pythonw.exe":
        python = python.with_name("python.exe")

    command = [str(python), "-X", "utf8", "-u", "-m", "mddm.run_merge",
               str(original), str(copies), "-o", str(output)]
    environment = os.environ.copy()
    package_parent = str(Path(__file__).resolve().parent.parent)
    python_path = environment.get("PYTHONPATH")
    environment["PYTHONPATH"] = os.pathsep.join(
        value for value in (package_parent, python_path) if value
    )
    try:
        with subprocess.Popen(
            command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace", env=environment,
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
        ) as process:
            for line in process.stdout:
                events.put(("log", line))
            exit_code = process.wait()
    except OSError as error:
        events.put(("log", f"启动合并失败: {error}\n"))
        exit_code = 1

    events.put(("done", exit_code))


class MergeWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("多部门文档自动合并系统")
        self.root.geometry("940x690")
        self.root.minsize(735, 590)
        self.root.configure(background=BACKGROUND)
        self.root.protocol("WM_DELETE_WINDOW", self.close)

        self.events = Queue()
        self.running = False
        self.pending_output = None
        self.last_suggestion = ""
        self.original = tk.StringVar(master=root)
        self.copies = tk.StringVar(master=root)
        self.output = tk.StringVar(master=root)
        self.form_controls = []
        self.path_entries = {}
        self.actions = []

        self._style()
        self._layout()

    def _style(self):
        family = "Microsoft YaHei UI" if os.name == "nt" else "TkDefaultFont"
        self.font = (family, 10)
        self.bold = (family, 10, "bold")
        self.style = ttk.Style(self.root)
        self.style.theme_use("clam")
        self.style.configure("TButton", font=self.font, padding=(14, 8),
                             foreground=INK, background="#F0F4FA", bordercolor=BORDER)
        self.style.map("TButton", background=[("active", "#E3ECF8")],
                       foreground=[("disabled", "#8290A2")])
        self.style.configure("Primary.TButton", font=self.bold, padding=(24, 10),
                             background=BLUE, foreground="white", bordercolor=BLUE)
        self.style.map("Primary.TButton",
                       background=[("disabled", "#AABAD3"), ("active", "#153E94")],
                       foreground=[("disabled", "white")])
        self.style.configure("Path.TEntry", padding=(9, 7), foreground=INK,
                             fieldbackground="#F9FBFE", bordercolor=BORDER)
        self.style.map("Path.TEntry", bordercolor=[("focus", BLUE)],
                       fieldbackground=[("disabled", "#EFF3F8")])
        self.style.configure("Merge.Horizontal.TProgressbar", troughcolor="#DCE6F3",
                             background=BLUE, borderwidth=0, thickness=5)

    def _layout(self):
        surface = tk.Frame(self.root, background=BACKGROUND, padx=30, pady=25)
        surface.pack(fill="both", expand=True)

        tk.Label(surface, text="合并 Word 文档", font=(self.font[0], 26, "bold"),
                 background=BACKGROUND, foreground=INK, anchor="w").pack(fill="x", pady=(0, 19))

        form = tk.Frame(surface, background=PANEL, highlightthickness=1,
                        highlightbackground=BORDER, padx=17, pady=13)
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)
        fields = (
            ("原稿文件", self.original, self._browse_original, "选择文件"),
            ("部门副本目录", self.copies, self._browse_copies, "选择目录"),
            ("合稿保存位置", self.output, self._browse_output, "另存为"),
        )
        for row, (label, variable, command, caption) in enumerate(fields):
            tk.Label(form, text=label, font=self.bold, background=PANEL, foreground=INK,
                     anchor="w").grid(row=row, column=0, sticky="w", padx=(0, 19), pady=5)
            entry = ttk.Entry(form, textvariable=variable, font=self.font, style="Path.TEntry", width=26)
            entry.grid(row=row, column=1, sticky="ew", pady=5)
            browse = ttk.Button(form, text=caption, width=9, command=command)
            browse.grid(row=row, column=2, sticky="e", padx=(11, 0), pady=5)
            self.form_controls.extend((entry, browse))
            self.path_entries[str(variable)] = entry

        controls = tk.Frame(surface, background=BACKGROUND)
        controls.pack(fill="x", pady=(14, 7))
        self.start_button = ttk.Button(controls, text="开始合并", style="Primary.TButton", command=self.start)
        self.start_button.pack(side="left")
        self.status = tk.Label(controls, text="等待开始", font=self.bold,
                               background=BACKGROUND, foreground=MUTED, anchor="w")
        self.status.pack(side="left", padx=13)
        self.progress = ttk.Progressbar(surface, mode="determinate", style="Merge.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(0, 15))

        tk.Label(surface, text="运行日志", font=self.bold, background=BACKGROUND, foreground=INK,
                 anchor="w").pack(fill="x", pady=(0, 7))
        console_frame = tk.Frame(surface, background=CONSOLE)
        console_frame.pack(fill="both", expand=True)
        self.console = tk.Text(console_frame, height=8, width=1, bg=CONSOLE, fg="#ECF3FF",
                               selectbackground="#3E628D", selectforeground="white", font=self.font,
                               borderwidth=0, highlightthickness=0, padx=14, pady=11,
                               wrap="word", state="disabled")
        scrollbar = ttk.Scrollbar(console_frame, orient="vertical", command=self.console.yview)
        self.console.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.console.pack(side="left", fill="both", expand=True)

        footer = tk.Frame(surface, background=BACKGROUND)
        footer.pack(fill="x", pady=(13, 0))
        for title in ("打开合稿", "打开冲突报告", "打开变更日志"):
            button = ttk.Button(footer, text=title, state="disabled")
            button.pack(side="left", padx=(0, 8))
            self.actions.append(button)

    def _set_path(self, variable, value):
        variable.set(value)
        self.path_entries[str(variable)].xview_moveto(1)

    def _suggest_output(self):
        if not self.original.get() or self.output.get() not in ("", self.last_suggestion):
            return
        original = Path(self.original.get())
        folder = original.parent
        if self.copies.get() and folder.resolve() == Path(self.copies.get()).resolve():
            folder = folder.parent
        self.last_suggestion = str(folder / f"{original.stem}_merged.docx")
        self._set_path(self.output, self.last_suggestion)

    def _browse_original(self):
        selected = filedialog.askopenfilename(parent=self.root, title="选择 Word 原稿", filetypes=DOCX_FILTER)
        if selected:
            self._set_path(self.original, selected)
            self._suggest_output()

    def _browse_copies(self):
        selected = filedialog.askdirectory(parent=self.root, title="选择部门副本所在目录")
        if selected:
            self._set_path(self.copies, selected)
            self._suggest_output()

    def _browse_output(self):
        current = Path(self.output.get()) if self.output.get() else Path("合稿.docx")
        selected = filedialog.asksaveasfilename(
            parent=self.root, title="选择合稿保存位置", filetypes=DOCX_FILTER,
            defaultextension=".docx", initialdir=str(current.parent), initialfile=current.name,
            confirmoverwrite=False,
        )
        if selected:
            self._set_path(self.output, selected)

    def start(self):
        if self.running:
            return
        try:
            paths = validate_paths(self.original.get(), self.copies.get(), self.output.get())
        except (ValueError, OSError, RuntimeError) as error:
            self.status.configure(text=str(error), foreground=RED)
            return

        output = paths[2]
        if any(path.exists() for path in result_paths(output)):
            confirmed = messagebox.askyesno("覆盖现有输出", "合稿或相关报告已存在，是否覆盖？",
                                            parent=self.root, default="no")
            if not confirmed:
                return

        self.running = True
        self.pending_output = output
        self._set_running_state(True)
        self._write_log("", clear=True)
        self.status.configure(text="正在合并…", foreground=BLUE)
        self.progress.configure(mode="indeterminate", value=0)
        self.progress.start(14)
        Thread(target=run_merge, args=(paths, self.events), daemon=True).start()
        self.root.after(60, self._poll_events)

    def _set_running_state(self, running):
        state = "disabled" if running else "normal"
        for control in [*self.form_controls, self.start_button]:
            control.configure(state=state)
        if running:
            for button in self.actions:
                button.configure(state="disabled")

    def _write_log(self, text, clear=False):
        self.console.configure(state="normal")
        if clear:
            self.console.delete("1.0", "end")
        if text:
            self.console.insert("end", text)
            self.console.see("end")
        self.console.configure(state="disabled")

    def _poll_events(self):
        lines = []
        exit_code = None
        for _ in range(400):
            try:
                kind, value = self.events.get_nowait()
            except Empty:
                break
            if kind == "log":
                lines.append(value)
            else:
                exit_code = value
        if lines:
            self._write_log("".join(lines))

        if exit_code is None:
            self.root.after(60, self._poll_events)
            return

        self.progress.stop()
        self._set_running_state(False)
        self.running = False
        output = self.pending_output
        succeeded = exit_code == 0 and output.is_file()
        self.progress.configure(mode="determinate", value=100 if succeeded else 0)
        if succeeded:
            self.status.configure(text="已生成合稿", foreground=GREEN)
            for button, path in zip(self.actions, result_paths(output)):
                if path.is_file():
                    button.configure(state="normal", command=lambda target=path: self._open_result(target))
        else:
            self.status.configure(text="合并失败，请查看运行日志", foreground=RED)

    def _open_result(self, path):
        try:
            if os.name == "nt":
                os.startfile(path)
            else:
                subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(path)])
        except OSError as error:
            messagebox.showerror("无法打开文件", str(error), parent=self.root)

    def close(self):
        if self.running:
            self.status.configure(text="合并正在进行，完成后可关闭窗口", foreground=BLUE)
        else:
            self.root.destroy()


def main():
    root = tk.Tk()
    MergeWindow(root)
    root.mainloop()


if __name__ == "__main__":
    main()
