"""バッジ判定ツールの GUI。Python を使わない職員でも操作できるようにする。"""
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from main import process


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('JCPP オープンバッジ 処理ツール')
        self.resizable(False, False)
        # Tkinter の画面は作業用スレッドから直接触れないので、キュー経由で受け渡す
        self.messages = queue.Queue()
        self._build_ui()

    def _build_ui(self):
        pad = {'padx': 12, 'pady': 6}

        tk.Label(self, text='入力ファイル（Excel）:', anchor='w').grid(row=0, column=0, sticky='w', **pad)
        self.input_var = tk.StringVar()
        tk.Entry(self, textvariable=self.input_var, width=48).grid(row=0, column=1, **pad)
        tk.Button(self, text='参照', command=self._pick_input).grid(row=0, column=2, **pad)

        tk.Label(self, text='出力ファイル（Excel）:', anchor='w').grid(row=1, column=0, sticky='w', **pad)
        self.output_var = tk.StringVar()
        tk.Entry(self, textvariable=self.output_var, width=48).grid(row=1, column=1, **pad)
        tk.Button(self, text='参照', command=self._pick_output).grid(row=1, column=2, **pad)

        self.run_btn = tk.Button(self, text='実行', width=20, bg='#4a90d9', fg='white',
                                 font=('', 12, 'bold'), command=self._run)
        self.run_btn.grid(row=2, column=0, columnspan=3, pady=10)

        self.progress = ttk.Progressbar(self, mode='indeterminate', length=420)
        self.progress.grid(row=3, column=0, columnspan=3, padx=12)

        self.log_text = tk.Text(self, height=8, width=60, state='disabled', bg='#f5f5f5', relief='sunken')
        self.log_text.grid(row=4, column=0, columnspan=3, padx=12, pady=8)

    def _pick_input(self):
        path = filedialog.askopenfilename(title='入力ファイルを選択', filetypes=[('Excel ファイル', '*.xlsx *.xls')])
        if path:
            self.input_var.set(path)
            if not self.output_var.get():
                self.output_var.set(path.rsplit('.', 1)[0] + '_output.xlsx')

    def _pick_output(self):
        path = filedialog.asksaveasfilename(title='出力ファイルを保存', defaultextension='.xlsx',
                                            filetypes=[('Excel ファイル', '*.xlsx')])
        if path:
            self.output_var.set(path)

    def _log(self, msg: str):
        self.log_text.configure(state='normal')
        self.log_text.insert('end', msg + '\n')
        self.log_text.see('end')
        self.log_text.configure(state='disabled')

    def _run(self):
        inp = self.input_var.get().strip()
        out = self.output_var.get().strip()
        if not inp or not out:
            messagebox.showwarning('確認', '入力・出力ファイルを指定してください。')
            return

        self.run_btn.configure(state='disabled')
        self.progress.start(10)

        def work():
            try:
                process(inp, out, log=lambda m: self.messages.put(('log', m)))
                self.messages.put(('done', out))
            except Exception as e:
                self.messages.put(('error', str(e)))

        threading.Thread(target=work, daemon=True).start()
        self.after(100, self._poll)

    def _poll(self):
        """作業用スレッドからのメッセージをメインスレッドで画面に反映する。"""
        while not self.messages.empty():
            kind, payload = self.messages.get()
            if kind == 'log':
                self._log(payload)
                continue
            self.progress.stop()
            self.run_btn.configure(state='normal')
            if kind == 'done':
                messagebox.showinfo('完了', f'出力完了:\n{payload}')
            else:
                self._log(f'エラー: {payload}')
                messagebox.showerror('エラー', payload)
            return
        self.after(100, self._poll)


if __name__ == '__main__':
    App().mainloop()
