# OPUS-Spectral-Data-Converter
This Python project provides a batch data conversion tool for Bruker OPUS files, which are commonly used in FTIR analysis.

It extracts raw spectral data (such as absorbance, single-channel signals) and exports them into structured `.xlsx` files with multiple sheets.

The tool was primarily developed for use in Room 202, Agricultural Sensing Department, Suyuan Building, with major contributions by Ke Wang.

Special thanks to Changbo Song and Lianglin Hao, whose priceless insights emerged during a legendary dinner of "酸萝卜鱼"(pickled radish fish).

## Script Descriptions

`opus_to_dpt_batch.py`  
Convert individual Bruker OPUS files into `.dpt` format using a reference template file. Outputs one `.dpt` file per OPUS file.

`opus_to_excel_batch.py`  
Batch extract spectral data from OPUS files and export to a timestamped Excel file (`.xlsx`) with multiple sheets, each corresponding to a spectral type (e.g., `AB`, `ScSm`).
