import os
import glob
from decimal import Decimal, ROUND_HALF_EVEN
from brukeropusreader import read_file
from pathlib import Path
import re

def convert_opus_to_dpt(input_opus_folder, output_dpt_folder, dpt_template_file):
    """
    Convert OPUS binary files to .dpt format using a provided X-axis template.
    All converted .dpt files will be saved directly to the output_dpt_folder.
    """

    # 确保输出文件夹存在
    Path(output_dpt_folder).mkdir(parents=True, exist_ok=True)

    # 读取模板 .dpt 文件中的 X 轴
    with open(dpt_template_file, "r", encoding="utf-8") as file:
        dpt_lines = file.readlines()
        x_values = [line.split(",")[0].strip() for line in dpt_lines]

    # 遍历输入目录中的所有 OPUS 文件
    opus_files = []
    pattern = re.compile(r'\.\d+$')  #匹配扩展名为 .任意长度的数字
    for root, _, files in os.walk(input_opus_folder):
        for file in files:
            if pattern.search(file):  # 可按需要调整扩展名过滤
                opus_files.append(os.path.join(root, file))

    if not opus_files:
        print("❗未找到任何 OPUS 文件")
        return

    print(f"🔍 共找到 {len(opus_files)} 个 OPUS 文件，开始转换...\n")

    for opus_path in opus_files:
        try:
            opus_data = read_file( opus_path )
            y_values = opus_data["ScSm"]
            # y_values = opus_data["AB"]

            # 构建输出文件名（只保留原文件名）
            file_name = os.path.basename(opus_path) + ".dpt"
            output_path = os.path.join(output_dpt_folder, file_name)

            # 写入 X,Y 数据
            with open(output_path, "w", encoding="utf-8") as out_file:
                for x, y in zip(x_values, y_values):
                    y_formatted = Decimal(str(y)).quantize(Decimal('.00001'), rounding=ROUND_HALF_EVEN)
                    out_file.write(f"{x},{y_formatted}\n")

            print(f"✅ 成功生成：{output_path}")

        except Exception as e:
            print(f"❌ 转换失败 {opus_path}，错误信息：{e}")

    print("\n🎉 所有文件转换完成！")

# 🔧 设置路径
if __name__ == "__main__":
    input_opus_folder = r"D:\Experiment data\实验数据\OPUS_DPT_数据转化\1、OPUS数据"
    output_dpt_folder = r"D:\Experiment data\实验数据\OPUS_DPT_数据转化\2、DPT数据"
    dpt_template_file = r"D:\Experiment data\实验数据\OPUS_DPT_数据转化\dpt_template\template.dpt"

    convert_opus_to_dpt(input_opus_folder, output_dpt_folder, dpt_template_file)
