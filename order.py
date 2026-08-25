import os
import shutil
import stat
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime, timedelta

def find_latest_file(pattern="宽带开通工单", ext=".xlsx"):
    files = []
    for f in os.listdir('.'):
        if f.startswith(pattern) and f.endswith(ext):
            files.append((f, os.path.getmtime(f)))
    if not files:
        return None
    files.sort(key=lambda x: x[1], reverse=True)
    return files[0][0]

def ensure_writable(filepath):
    try:
        current_mode = os.stat(filepath).st_mode
        os.chmod(filepath, current_mode | stat.S_IWUSR | stat.S_IRUSR)
    except Exception as e:
        print(f"设置文件权限警告（非关键）：{e}")

def clean_town(value):
    if pd.isna(value):
        return value
    value_str = str(value)
    if value_str.endswith("镇"):
        return value_str[:-1]
    return value_str

def main():
    # 1. 定位最新源文件
    src_file = find_latest_file("宽带开通工单", ".xlsx")
    if src_file is None:
        print("错误：未找到以'宽带开通工单'开头的.xlsx文件。")
        return
    print(f"使用源文件：{src_file}")

    # 2. 读取源数据（IVR未接通数据）
    try:
        df_src = pd.read_excel(src_file, sheet_name="IVR未接通数据", dtype=str)
    except Exception as e:
        print(f"读取源文件失败：{e}")
        return

    # 3. 过滤业务范围
    allowed_business = ['FTTR开通', '宽带开通', '全屋WIFI基础产品开通']
    if '业务范围' in df_src.columns:
        df_src = df_src[df_src['业务范围'].isin(allowed_business)]
        print(f"过滤后剩余 {len(df_src)} 条记录")
    else:
        print("警告：源数据中无'业务范围'列，将不过滤。")

    # 4. 映射关系（导用sheet）
    mapping = {
        "用户号码": "客户号码",
        "呼叫时间": "系统产品名称",
        "业务范围": "业务范围",
        "分公司": "分公司",
        "镇区": "镇区",
        "归档日期": "归档日期",
        "地址": "地址",
    }

    # 5. 读取模板列名
    template_file = "导数模板.xlsx"
    try:
        df_template = pd.read_excel(template_file, sheet_name="导用", nrows=0)
        target_columns = df_template.columns.tolist()
    except Exception as e:
        print(f"读取模板失败：{e}")
        return

    # 6. 构建导用DataFrame
    df_out = pd.DataFrame(columns=target_columns)
    for src_col, tgt_col in mapping.items():
        if src_col in df_src.columns and tgt_col in target_columns:
            df_out[tgt_col] = df_src[src_col]
    df_out = df_out[target_columns]

    # 7. 处理镇区列（去掉末尾“镇”）
    if "镇区" in df_out.columns:
        df_out["镇区"] = df_out["镇区"].apply(clean_town)

    # 8. 生成镇区汇总数据
    if "镇区" in df_out.columns:
        town_counts = df_out["镇区"].value_counts().reset_index()
        town_counts.columns = ["镇区", "数量"]
    else:
        print("错误：导用数据中无'镇区'列，无法生成汇总。")
        return

    # 9. 日期计算（今天、明天、后天，共3天）
    today = datetime.now().date()
    start_date = today.strftime("%m%d")                       # 如 0825
    end_date = (today + timedelta(days=2)).strftime("%m%d")   # 如 0827
    date_range = f"【{start_date}-{end_date}】新装宽带客户回访"
    today_display = f"{today.month}月{today.day}日"          # 如 8月25日（无前导零）

    # 10. 构建汇总DataFrame，并追加汇总行（只填充数量）
    df_summary = pd.DataFrame({
        "日期范围": [date_range] * len(town_counts),
        "日期": [today_display] * len(town_counts),
        "镇区": town_counts["镇区"],
        "数量": town_counts["数量"]
    })
    total_count = town_counts["数量"].sum()
    summary_row = pd.DataFrame({
        "日期范围": [""],
        "日期": [""],
        "镇区": [""],
        "数量": [total_count]
    })
    df_summary = pd.concat([df_summary, summary_row], ignore_index=True)

    # 11. 输出主文件路径
    output_file = "导数模板_处理后.xlsx"

    # 12. 删除旧的主文件（若存在）
    if os.path.exists(output_file):
        try:
            os.remove(output_file)
            print(f"已删除旧文件：{output_file}")
        except PermissionError:
            print(f"错误：文件 {output_file} 正在被占用，请关闭后重新运行。")
            return
        except Exception as e:
            print(f"删除旧文件失败：{e}")
            return

    # 13. 复制模板（保留所有样式）
    try:
        shutil.copy2(template_file, output_file)
    except Exception as e:
        print(f"复制模板失败：{e}")
        return

    # 14. 确保新文件可写
    ensure_writable(output_file)

    # 15. 写入数据到“导用”和“镇区汇总”
    try:
        wb = load_workbook(output_file)
        ws = wb["导用"]
        ws.delete_rows(2, ws.max_row)  # 清空旧数据行

        for r_idx, row in enumerate(dataframe_to_rows(df_out, index=False, header=False), start=2):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        ws_summary = wb.create_sheet("镇区汇总")
        for c_idx, col_name in enumerate(df_summary.columns, start=1):
            ws_summary.cell(row=1, column=c_idx, value=col_name)
        for r_idx, row in enumerate(dataframe_to_rows(df_summary, index=False, header=False), start=2):
            for c_idx, value in enumerate(row, start=1):
                ws_summary.cell(row=r_idx, column=c_idx, value=value)

        wb.save(output_file)
        print(f"主文件已生成：{output_file}（包含导用和镇区汇总）")
    except Exception as e:
        print(f"写入主文件失败：{e}")
        return

    # ========== 16. 按镇区分拆导出 ==========
    folder_name = today.strftime("%m%d")  # 如 "0825"
    try:
        if os.path.exists(folder_name):
            shutil.rmtree(folder_name)  # 删除旧文件夹（若有）
        os.makedirs(folder_name)
        print(f"创建文件夹：{folder_name}")
    except Exception as e:
        print(f"创建文件夹失败：{e}")
        return

    # 获取所有镇区
    towns = df_out["镇区"].dropna().unique()
    if len(towns) == 0:
        print("没有有效镇区数据，跳过拆分。")
        return

    for town in towns:
        df_town = df_out[df_out["镇区"] == town].copy()
        # 如果镇区名为空，跳过
        if pd.isna(town) or town == "":
            continue
        # 安全处理文件名（去除非法字符）
        safe_name = str(town).replace("/", "_").replace("\\", "_").replace(":", "_")
        file_path = os.path.join(folder_name, f"{safe_name}.xlsx")
        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df_town.to_excel(writer, sheet_name="导用", index=False)
            print(f"  已生成：{file_path}（{len(df_town)}条记录）")
        except Exception as e:
            print(f"  写入 {file_path} 失败：{e}")

    print("所有镇区分拆文件已生成完毕。")

if __name__ == "__main__":
    main()