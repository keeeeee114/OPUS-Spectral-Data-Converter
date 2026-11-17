import os
import re
from decimal import Decimal, ROUND_HALF_EVEN
from brukeropusreader import read_file
import pandas as pd
from datetime import datetime
"""
未Debug，目前只测试AB，加后再放github
"""
def auto_convert_opus_to_excel(
    input_opus_folder,
    output_excel_folder,
    fields_to_extract=("AB", "ScSm")
):
    opus_files = []
    pattern = re.compile(r"\.\d+$")
    for root, _, files in os.walk(input_opus_folder):
        for file in files:
            if pattern.search(file):
                opus_files.append(os.path.join(root, file))

    if not opus_files:
        print("没有找到 OPUS 文件")
        return

    print(f" 共找到 {len(opus_files)} 个 OPUS 文件，开始处理...\n")
    failed_files = []

    # 存储每个 field 的数据
    all_data = {field: {} for field in fields_to_extract}
    # 存储每个 field 的 X 轴
    x_axes = {field: None for field in fields_to_extract}

    for opus_path in opus_files:
        try:
            data = read_file(opus_path)
            sample_name = os.path.basename(opus_path)

            for field in fields_to_extract:
                y_values = data.get(field)
                if y_values is None:
                    print(f"{sample_name} 中未找到字段 {field}，跳过")
                    continue

                # 获取该 field 对应的波数
                x_vals_raw = data.get_range(field)

                # 若该 field 的波数尚未存储，则存入
                if x_axes[field] is None:
                    # 保留高精度（因光谱仪读数通常超5位小数）
                    x_axes[field] = [Decimal(str(x)).quantize(Decimal('.00001'), rounding=ROUND_HALF_EVEN) for x in x_vals_raw]

                # 处理多通道数据
                if isinstance(y_values[0], (list, tuple)):
                    y_vals_raw = [y[0] for y in y_values]
                else:
                    y_vals_raw = y_values

                # 保留5位小数
                y_vals = [Decimal(str(y)).quantize(Decimal('.00001'), rounding=ROUND_HALF_EVEN) for y in y_vals_raw]

                # 长度匹配检查
                if len(y_vals) != len(x_axes[field]):
                    print(f"️ {sample_name} | {field}: 长度不匹配 (X={len(x_axes[field])}, Y={len(y_vals)})")
                else:
                    all_data[field][sample_name] = y_vals

        except Exception as e:
            print(f" 文件处理失败 {opus_path}：{e}")
            failed_files.append(opus_path)

    print("\n数据提取完毕，开始写入 Excel...\n")
    print("成功提取的数据样本数（按字段统计）：")
    for field, samples in all_data.items():
        print(f" - {field}: {len(samples)}")

    # 如果没有任何数据，报错退出
    if all(len(samples) == 0 for samples in all_data.values()):
        print("\n错误：没有成功读取任何数据，无法生成 Excel 文件")
        if failed_files:
            print("\n检查方式：")
            print("  1. 手动测试：print(opus_data.keys()) 查看文件中有哪些字段")
            print("  2. 尝试使用 fields_to_extract = ('AB', 'ScSm', 'ScRf')")
            print("  3. 检查是否有文件的字段名称不一样")
        return

    # 生成 Excel
    timestamp = datetime.now().strftime('%y%m%d_%H%M%S')
    output_excel_path = os.path.join(output_excel_folder, f'光谱数据_{timestamp}.xlsx')

    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
        for field, samples in all_data.items():
            if not samples or x_axes[field] is None:
                continue

            df = pd.DataFrame(samples)
            x_vals = x_axes[field]
            if len(x_vals) != len(df):
                print(f"写入前长度不匹配，截断至 {min(len(x_vals), len(df))}")
                min_len = min(len(x_vals), len(df))
                x_vals = x_vals[:min_len]
                df = df.iloc[:min_len, :]
            df.insert(0, "X", x_vals)
            df.to_excel(writer, sheet_name=field, index=False)
            print(f"写入 sheet：{field}（{len(samples)} 个样本）")

    print(f"\n 所有数据已写入 Excel 文件：{output_excel_path}")

if __name__ == "__main__":
    input_opus_folder = r"D:\新建文件夹"
    output_excel_folder = r"D:\新建文件夹"

    auto_convert_opus_to_excel(
        input_opus_folder,
        output_excel_folder,
        fields_to_extract=("AB", "ScSm")
    )
