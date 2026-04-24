#!/usr/bin/env python3
"""
Excel批号合并工具 - GUI版本
根据批号将多行数据合并成一行，对指定字段求和
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from pathlib import Path
import threading


def merge_by_batch(input_file: str, output_file: str = None):
    """根据批号合并Excel中的多行数据"""
    df = pd.read_excel(input_file)

    keep_fields = ['序号', '商品编号', '商品名称', '商品规格', '剂型', '件包装数', '单位',
                  '生产企业', '批号', '生产日期', '有效期至', '存储条件']

    sum_fields = ['库管数量', '件数', '零散数量']

    for field in keep_fields + sum_fields:
        if field not in df.columns:
            raise ValueError(f"找不到字段: {field}, 可用字段: {df.columns.tolist()}")

    agg_dict = {field: 'sum' for field in sum_fields}
    for field in keep_fields:
        if field not in agg_dict:
            agg_dict[field] = 'first'

    merged_df = df.groupby(['商品编号', '批号'], as_index=False).agg(agg_dict)

    carry = merged_df['零散数量'] // merged_df['件包装数']
    merged_df['件数'] = merged_df['件数'] + carry
    merged_df['零散数量'] = merged_df['零散数量'] % merged_df['件包装数']

    merged_df = merged_df[keep_fields + sum_fields]
    merged_df['序号'] = range(1, len(merged_df) + 1)

    if output_file is None:
        input_path = Path(input_file)
        output_file = input_path.parent / f"{input_path.stem}_merged.xlsx"

    merged_df.to_excel(output_file, index=False)
    return output_file, len(df), len(merged_df)


class MergeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel批号合并工具")
        self.root.geometry("500x180")
        self.root.resizable(False, False)

        # 设置样式
        style = ttk.Style()
        style.configure("TButton", padding=6)
        style.configure("TLabel", font=("微软雅黑", 10))

        # 主框架
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 选择文件
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=5)
        ttk.Label(file_frame, text="选择文件:").pack(side=tk.LEFT)
        self.file_var = tk.StringVar()
        self.file_entry = ttk.Entry(file_frame, textvariable=self.file_var, width=35)
        self.file_entry.pack(side=tk.LEFT, padx=5)
        ttk.Button(file_frame, text="浏览", command=self.browse_file).pack(side=tk.LEFT)

        # 运行按钮
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=20)
        self.run_btn = ttk.Button(btn_frame, text="开始合并", command=self.run_merge, width=15)
        self.run_btn.pack()

        # 状态显示
        self.status_label = ttk.Label(main_frame, text="就绪", foreground="gray")
        self.status_label.pack(pady=5)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xls *.xlsx *.xlsm"), ("所有文件", "*.*")]
        )
        if file_path:
            self.file_var.set(file_path)

    def run_merge(self):
        input_file = self.file_var.get()
        if not input_file:
            messagebox.showwarning("警告", "请选择要合并的Excel文件")
            return

        file_path = Path(input_file)
        if not file_path.exists():
            messagebox.showerror("错误", f"文件不存在: {input_file}")
            return

        # 禁用按钮
        self.run_btn.config(state=tk.DISABLED)
        self.status_label.config(text="处理中...", foreground="blue")

        # 在线程中执行
        def worker():
            try:
                output_file, original_count, merged_count = merge_by_batch(input_file)
                self.root.after(0, lambda out=output_file, orig=original_count, merg=merged_count: self.on_complete(True, out, orig, merg))
            except Exception as e:
                self.root.after(0, lambda err=str(e): self.on_complete(False, err))

        threading.Thread(target=worker, daemon=True).start()

    def on_complete(self, success, output_file_or_error, original_count=None, merged_count=None):
        self.run_btn.config(state=tk.NORMAL)
        if success:
            self.status_label.config(text=f"完成！原始{original_count}行 → 合并{merged_count}行", foreground="green")
            messagebox.showinfo("成功", f"合并完成！\n输出文件: {output_file_or_error}")
        else:
            self.status_label.config(text="失败", foreground="red")
            messagebox.showerror("错误", f"合并失败:\n{output_file_or_error}")


def main():
    root = tk.Tk()
    app = MergeApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()