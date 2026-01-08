import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def compare_files(file1_path, file2_path, output_path='report.xlsx'):
    # ========= 读取文件1：完整符文表（保持原始结构） =========
    df1 = pd.read_excel(file1_path, sheet_name='Sheet1')

    # A列作为符文名称（主键）
    rune_col = df1.columns[0]

    # 清洗 A 列
    df1 = df1.dropna(subset=[rune_col])
    df1[rune_col] = df1[rune_col].astype(str).str.strip()

    # ========= 读取文件2：已拥有符文（仅用于判断） =========
    df2 = pd.read_excel(file2_path, sheet_name='Sheet1', header=None)

    owned_set = set()
    for col in df2.columns:
        owned_set.update(
            df2[col]
            .dropna()
            .astype(str)
            .str.strip()
            .tolist()
        )

    # ========= 核心逻辑：只筛选，不改结构 =========
    missing_df = df1[~df1[rune_col].isin(owned_set)]

    # ========= 写入 Excel（格式与文件1一致） =========
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        missing_df.to_excel(writer, index=False, sheet_name='Sheet1')

    print(f'对比报告已生成：{output_path}')
    print(f'文件1 总数量：{len(df1)}')
    print(f'文件2 已有数量：{len(owned_set)}')
    print(f'缺少符文数：{len(missing_df)}')

if __name__ == "__main__":
    Tk().withdraw()

    print("请选择【参照文件】（文件1）...")
    file1 = askopenfilename(
        title="选择参照文件",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not file1:
        print("未选择文件，程序退出。")
        exit()

    print("请选择【对照文件】（文件2）...")
    file2 = askopenfilename(
        title="选择对照文件",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not file2:
        print("未选择文件，程序退出。")
        exit()

    compare_files(file1, file2)
