#!/usr/bin/env python3
"""
Excel批号合并工具
根据批号将多行数据合并成一行，对指定字段求和
"""

import sys
import pandas as pd
from pathlib import Path


def merge_by_batch(input_file: str, output_file: str = None):
    """
    根据批号合并Excel中的多行数据

    Args:
        input_file: 输入的xls/xlsx文件名
        output_file: 输出的xlsx文件名，默认为输入文件加上'_merged'后缀
    """
    # 读取Excel文件
    df = pd.read_excel(input_file)

    # 定义需要保留的字段（不去重，取第一条记录的值）
    keep_fields = ['序号', '商品编号', '商品名称', '商品规格', '剂型', '件包装数', '单位',
                  '生产企业', '批号', '生产日期', '有效期至', '存储条件']

    # 定义需要求和的字段
    sum_fields = ['库管数量', '件数', '零散数量']

    # 验证字段是否存在
    for field in keep_fields + sum_fields:
        if field not in df.columns:
            raise ValueError(f"找不到字段: {field}, 可用字段: {df.columns.tolist()}")

    # 按批号分组，求和并保留第一条记录的保留字段
    # 首先，对需要求和的字段进行分组求和
    agg_dict = {field: 'sum' for field in sum_fields}
    for field in keep_fields:
        if field not in agg_dict:
            agg_dict[field] = 'first'  # 保留第一条的值

    # 执行分组聚合（按商品编号+批号分组）
    merged_df = df.groupby(['商品编号', '批号'], as_index=False).agg(agg_dict)

    # 对零散数量进行进位处理：当零散数量 >= 件包装数时，进位到件数
    carry = merged_df['零散数量'] // merged_df['件包装数']
    merged_df['件数'] = merged_df['件数'] + carry
    merged_df['零散数量'] = merged_df['零散数量'] % merged_df['件包装数']

    # 重新排列列顺序
    merged_df = merged_df[keep_fields + sum_fields]

    # 重新生成序号
    merged_df['序号'] = range(1, len(merged_df) + 1)

    # 确定输出文件名
    if output_file is None:
        input_path = Path(input_file)
        output_file = input_path.parent / f"{input_path.stem}_merged.xlsx"

    # 保存结果
    merged_df.to_excel(output_file, index=False)
    print(f"合并完成: {output_file}")
    print(f"原始行数: {len(df)}, 合并后行数: {len(merged_df)}")


def main():
    if len(sys.argv) < 2:
        print("用法: python merge_by_batch.py <输入文件> [输出文件]")
        print("示例: python merge_by_batch.py data.xls")
        print("示例: python merge_by_batch.py data.xls output.xlsx")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None

    # 检查输入文件是否存在
    if not Path(input_file).exists():
        print(f"错误: 文件不存在: {input_file}")
        sys.exit(1)

    try:
        merge_by_batch(input_file, output_file)
    except Exception as e:
        print(f"错误: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()