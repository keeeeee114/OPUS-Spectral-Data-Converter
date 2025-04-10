import os
import re
from decimal import Decimal, ROUND_HALF_EVEN
from brukeropusreader import read_file
import pandas as pd
from datetime import datetime

# Environment Info
# Python 3.9
# numpy==1.22.4        # from conda-forge
# pandas==1.5.3         # from conda-forge
# brukeropusreader==1.3.4

def auto_convert_opus_to_excel(
    input_opus_folder,
    output_excel_folder,
    dpt_template_file,
    fields_to_extract=("AB", "ScSm", "RT", "TR") #可以自行扩展
    # 吸光度光谱      AB == Absorbance
    # 样本的原始光谱   ScSm == Single channel sample
    # 反射率光谱      RT == Reflectance
    # 透射率光谱      TR == Transmission
):
    # 初始化字典：每个字段一个表
    field_dataframes = {field: {} for field in fields_to_extract}
    x_axes = {}

    # 匹配扩展名为数字的文件（.0, .0001等）
    opus_files = []
    pattern = re.compile(r"\.\d+$")
    for root, _, files in os.walk(input_opus_folder):
        for file in files:
            if pattern.search(file):
                opus_files.append(os.path.join(root, file))

    if not opus_files:
        print("❗ 没有找到 OPUS 文件")
        return

    print(f"📂 共找到 {len(opus_files)} 个 OPUS 文件，开始处理...\n")

    for opus_path in opus_files:
        try:
            data = read_file(opus_path)
            sample_name = os.path.basename(opus_path)

            for field in fields_to_extract:
                y_values = data.get(field)
                if y_values is None:
                    continue  # 此谱图未包含该字段

                # 读取模板 .dpt 文件中的 X 轴
                with open(dpt_template_file, "r", encoding="utf-8") as file:
                    dpt_lines = file.readlines()
                    x_values = [line.strip().split()[0].strip() for line in dpt_lines]

                # 处理多通道情况，只取第一个通道（或扩展多列也可）  保留小数点后5位
                if isinstance(y_values[0], (list, tuple)):
                    y_values = [Decimal(str(y[0])).quantize(Decimal('.00001'), rounding=ROUND_HALF_EVEN) for y in y_values]
                    #isinstance(y_values[0], (list, tuple)) 判断数据是否为多通道
                    #取ScSm第一列 ，Decimal.quantize保留5位小数，ROUND_HALF_EVEN 四舍六入 五成双
                else:
                    y_values = [Decimal(str(y)).quantize(Decimal('.00001'), rounding=ROUND_HALF_EVEN) for y in y_values]

                # 存入 field 的列
                field_dataframes[field][sample_name] = y_values
                # 存一次 X
                x_axes[field] = x_values

        except Exception as e:
            print(f"❌ 文件处理失败 {opus_path}：{e}")

    print("\n✅ 数据提取完毕，开始写入 Excel...\n")

    #生成时间戳
    timestamp = datetime.now().strftime('%y%m%d_%H%M%S')
    output_excel_path = os.path.join(output_excel_folder, f'光谱数据_{timestamp}.xlsx')

    # 写入 Excel 多个 Sheet
    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
        for field, samples in field_dataframes.items():
            if not samples:
                continue

            df = pd.DataFrame(samples)
            df.insert(0, "X", x_axes[field])
            df.to_excel(writer, sheet_name=field, index=False)

            print(f"📄 写入 sheet：{field}（{len(samples)} 个样本）")

    print(f"\n🎉 所有数据已写入 Excel 文件：{output_excel_path}")

# 示例调用
if __name__ == "__main__":
    input_opus_folder = r"D:\Experiment data\实验数据\OPUS_DPT_数据转化\1、OPUS数据"
    output_excel_folder = r"D:\Experiment data\实验数据\OPUS_DPT_数据转化"
    dpt_template_file = r"D:\Experiment data\实验数据\OPUS_DPT_数据转化\dpt_template\template.dpt"

    auto_convert_opus_to_excel(
        input_opus_folder,
        output_excel_folder,
        dpt_template_file,
        fields_to_extract=("AB", "ScSm")  # 可选其他字段
    )