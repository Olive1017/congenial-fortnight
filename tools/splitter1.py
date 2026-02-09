import os
from copy import copy
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.styles import Font, Alignment

# ========= 关键字 =========
HEADER_END_KEYS = ["ITEM NO.", "DESCRIPTION"]
COMPANY_KEY_1 = "P&G"
COMPANY_KEY_2 = "PROCTER & GAMBLE (GUANGZHOU) LTD."


def _row_text(row):
    return " ".join(str(c.value).strip() for c in row if c.value is not None)


def _is_header_start(row1_text, row2_text):
    has_pg = COMPANY_KEY_1 in row1_text
    has_full_name = COMPANY_KEY_2 in row2_text
    return has_pg and has_full_name


def _is_header_end(text):
    return all(k in text for k in HEADER_END_KEYS)


# 封装一个样式复制函数，确保所有属性都被保留
def copy_cell_style(source_cell, target_cell):
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = source_cell.number_format
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)


def split_excel_by_row(input_path, output_prefix, split_size=30):
    def _get_new_col_idx(old_idx):
        if old_idx <= 3: return old_idx
        return old_idx + 1

    wb = load_workbook(input_path)
    ws = wb.active
    max_col = ws.max_column

    # ===== 1. 找表头范围 =====
    header_starts = []
    header_ends = []
    rows_list = list(ws.iter_rows(min_row=2))

    for i, row in enumerate(rows_list):
        row_idx = i + 2
        text = _row_text(row)
        if _is_header_end(text):
            header_ends.append(row_idx)
        if i < len(rows_list) - 1:
            if _is_header_start(text, _row_text(rows_list[i + 1])):
                header_starts.append(row_idx)

    if not header_starts or not header_ends:
        raise ValueError("未找到有效表头")

    first_table_header_end = header_ends[0]

    # 修正表头高度（合并单元格检测）
    max_merge_row = first_table_header_end
    for merged in ws.merged_cells.ranges:
        if merged.min_row <= first_table_header_end <= merged.max_row:
            if merged.max_row > max_merge_row:
                max_merge_row = merged.max_row
    first_table_header_end = max_merge_row
    header_ends[0] = max_merge_row

    # ===== 2. 确定表格拆分范围 =====
    tables = []
    for i, header_start in enumerate(header_starts):
        header_end = header_ends[i] if i < len(header_ends) else header_start + 1
        table_start = 1 if i == 0 else header_start
        table_end = header_starts[i + 1] - 2 if i < len(header_starts) - 1 else ws.max_row
        tables.append({'start': table_start, 'end': table_end, 'header_start': header_start, 'header_end': header_end})

    # ===== 3. 准备第一张表头行号 =====
    first_table_header_row_nums = []
    for row_num in range(tables[0]['start'], first_table_header_end + 1):
        first_table_header_row_nums.append(row_num)

    # 找 ITEM NO 列和最后数值
    first_table_last_item_no = None
    first_table_item_col = None

    # (省略部分辅助查找逻辑，保持原样)
    for r_search in range(tables[0]['header_start'], tables[0]['header_end'] + 1):
        for c_idx, cell in enumerate(ws[r_search], 1):
            if cell.value and "ITEM NO" in str(cell.value):
                first_table_item_col = c_idx;
                break
        if first_table_item_col: break

    if first_table_item_col:
        data_start = tables[0]['header_end'] + 1
        footer_start = None
        for r in range(data_start, tables[0]['end'] + 1):
            if "TOTAL" in _row_text(ws[r]): footer_start = r; break
        end_row = footer_start - 1 if footer_start else tables[0]['end']
        for r in range(end_row, data_start - 1, -1):
            try:
                val = ws.cell(r, first_table_item_col).value
                if val is not None: first_table_last_item_no = int(val); break
            except:
                continue

    output_files = []
    output_dir = os.path.dirname(output_prefix)
    if output_dir: os.makedirs(output_dir, exist_ok=True)

    for idx, table_info in enumerate(tables, 1):
        new_wb = Workbook()
        new_ws = new_wb.active
        new_ws.page_margins = PageMargins(left=0, right=0, top=0, bottom=0, header=0.28, footer=0.12)

        # ===== 强制锁定打印效果 =====
        new_ws.page_setup.paperSize = new_ws.PAPERSIZE_A4  # 强制设为 A4 纸

        # 核心设置：强制将所有列缩放到一页宽
        # 这样即使对方打印机驱动有点偏差，Excel 也会自动微调比例让它刚好填满横向
        new_ws.sheet_properties.pageSetUpPr.fitToPage = True
        new_ws.page_setup.fitToHeight = 0  # 高度不限（随数据多少自动分页）
        new_ws.page_setup.fitToWidth = 1  # 宽度强制为 1 页

        # 让页面在打印时水平居中
        new_ws.print_options.horizontalCentered = True

        # 确定原始表头的行号列表
        raw_header_rows = list(first_table_header_row_nums)

        # 记录原始的第二行行号（如果存在），用于稍后提取 P&G 的样式
        original_row2_idx = None
        if len(raw_header_rows) >= 2:
            original_row2_idx = raw_header_rows[1]
            # 我们在新表中跳过第二行
            rows_to_write = [raw_header_rows[0]] + raw_header_rows[2:]
        else:
            rows_to_write = raw_header_rows

        # 压缩空白行逻辑
        compressed_header_rows = []
        prev_blank = False

        def is_blank_old_row(ws_obj, r_idx):
            return all(ws_obj.cell(r_idx, c).value in (None, "") for c in range(1, ws_obj.max_column + 1))

        for r in rows_to_write:
            if is_blank_old_row(ws, r):
                if not prev_blank: compressed_header_rows.append(r)
                prev_blank = True
            else:
                compressed_header_rows.append(r)
                prev_blank = False
        rows_to_write = compressed_header_rows

        # 确定数据起始行
        if idx == 1:
            data_start_old_row = table_info['header_end'] + 1
        else:
            data_start_old_row = table_info['header_start'] + 1
            if first_table_last_item_no is not None and first_table_item_col:
                exp = first_table_last_item_no + 1
                for r in range(table_info['header_start'], table_info['end'] + 1):
                    try:
                        if ws.cell(r, first_table_item_col).value == exp:
                            data_start_old_row = r;
                            break
                    except:
                        continue
        rows_to_write.extend(range(data_start_old_row, table_info['end'] + 1))

        row_map = {}
        new_r = 1

        # ===== 逐行写入（样式修复版） =====
        for old_r in rows_to_write:
            if old_r in ws.row_dimensions:
                new_ws.row_dimensions[new_r].height = ws.row_dimensions[old_r].height

            # 判断当前是否是新表的第一行
            is_new_first_row = (new_r == 1)

            for c_idx in range(1, max_col + 1):
                new_c = _get_new_col_idx(c_idx)

                # 默认源单元格
                source_cell = ws.cell(old_r, c_idx)

                # [关键修复]：如果是新表第一行，且位于 A-C 列 (P&G 区域)
                # 我们需要智能判断是取 Row 1 还是 Row 2 的内容和样式
                if is_new_first_row and new_c <= 3:
                    # 如果当前 Row 1 对应位置为空，且我们知道有 Row 2
                    if not source_cell.value and original_row2_idx:
                        # 尝试从 Row 2 取
                        cell_row2 = ws.cell(original_row2_idx, c_idx)
                        if cell_row2.value:
                            source_cell = cell_row2  # !!! 切换源单元格为 Row 2

                # 写入值
                new_cell = new_ws.cell(row=new_r, column=new_c, value=source_cell.value)

                # 复制样式 (使用封装函数)
                copy_cell_style(source_cell, new_cell)

                # 复制列宽
                l_old = get_column_letter(c_idx)
                l_new = get_column_letter(new_c)
                if l_old in ws.column_dimensions:
                    new_ws.column_dimensions[l_new].width = ws.column_dimensions[l_old].width

            # 补齐 D 列宽度
            if 'C' in ws.column_dimensions:
                new_ws.column_dimensions['D'].width = ws.column_dimensions['C'].width

            # -----------------------------------------------------
            # 第一行特殊处理：右侧标题合并 (保留样式版)
            # -----------------------------------------------------
            if is_new_first_row:
                # 1. 左侧 P&G 合并
                new_ws.merge_cells(start_row=1, end_row=1, start_column=1, end_column=3)

                # 2. 右侧长标题合并 (D列以后)
                start_col = 4
                end_col = new_ws.max_column

                # 收集文本
                parts = []
                for c in range(start_col, end_col + 1):
                    v = new_ws.cell(new_r, c).value
                    if v: parts.append(str(v).strip())

                if parts:
                    merged_text = "  ".join(parts)
                    target_cell = new_ws.cell(new_r, start_col)

                    # 在覆盖值之前，确保 target_cell 拥有正确的样式
                    # 通常 D列是空白的，样式可能在后面的列里。
                    # 我们找到第一个有值的列作为样式源
                    style_source_col = start_col
                    for c in range(start_col, end_col + 1):
                        if new_ws.cell(new_r, c).value:
                            style_source_col = c;
                            break

                    # 复制该列的样式到 D 列 (target_cell)
                    copy_cell_style(new_ws.cell(new_r, style_source_col), target_cell)

                    target_cell.value = merged_text

                    # 强制右对齐
                    target_cell.alignment = Alignment(horizontal="right", vertical="center")

                # 清空 D 列之后的内容防止重叠
                for c in range(start_col + 1, end_col + 1):
                    new_ws.cell(new_r, c).value = None

                new_ws.merge_cells(start_row=new_r, end_row=new_r, start_column=start_col, end_column=end_col)

            row_map[old_r] = new_r
            new_r += 1

        # ===== 处理原表合并单元格 =====
        for merged in ws.merged_cells.ranges:
            if merged.min_row in row_map and merged.max_row in row_map:
                new_min_c = _get_new_col_idx(merged.min_col)
                new_max_c = _get_new_col_idx(merged.max_col)
                if merged.min_col == 2 and merged.max_col == 3: new_max_c = 4

                # 避开我们已经处理过的第一行
                if not (row_map[merged.min_row] == 1):
                    new_ws.merge_cells(start_row=row_map[merged.min_row], end_row=row_map[merged.max_row],
                                       start_column=new_min_c, end_column=new_max_c)

        # ===== [关键修复] 字体放大逻辑 =====
        # 使用 copy() 而不是 Font() 构造函数，以保留颜色
        FONT_DELTA = 9
        for r in range(1, 2):  # 只处理第一行
            for c in range(1, new_ws.max_column + 1):
                cell = new_ws.cell(row=r, column=c)
                if cell.value:
                    if cell.font:
                        new_font = copy(cell.font)
                        # 安全地增加大小
                        new_font.size = (new_font.size if new_font.size else 11) + FONT_DELTA
                        new_font.bold = True  # 确保加粗
                        cell.font = new_font

        # 行高设置
        new_ws.row_dimensions[1].height = 36

        def find_data_start_row(ws):
            for r in range(1, ws.max_row + 1):
                row_text = " ".join(str(c.value) for c in ws[r] if c.value)
                if "ITEM NO" in row_text:
                    return r + 1
            raise ValueError("未找到 ITEM NO 行")

        data_start_row = find_data_start_row(new_ws)
        item_col_new = None
        for r_h in range(1, find_data_start_row(new_ws) + 1):
            for c_idx in range(1, new_ws.max_column + 1):
                cell = new_ws.cell(r_h, c_idx)
                if cell.value and "ITEM NO" in str(cell.value):
                    item_col_new = c_idx
                    break
            if item_col_new: break

        if item_col_new:
            num = 1
            footer_start_new_row = None
            for r in range(new_ws.max_row, data_start_row, -1):
                row_txt = "".join([str(c.value or "") for c in new_ws[r]])
                if "TOTAL" in row_txt:
                    footer_start_new_row = r
                    break

            data_end_row = footer_start_new_row - 1 if footer_start_new_row else new_ws.max_row

            for r in range(data_start_row, data_end_row + 1):
                v = new_ws.cell(r, item_col_new).value
                if v is not None:
                    new_ws.cell(r, item_col_new).value = num
                    num += 1

                # =========================================================
                # [精准版] 表尾 TOTAL DAP 及其右侧数值加粗逻辑
                # =========================================================
                try:
                    from openpyxl.styles import Font
                    found_dap = False
                    # 从最后 15 行向上扫描（表尾区域）
                    for r in range(new_ws.max_row, max(1, new_ws.max_row - 15), -1):
                        for c in range(1, new_ws.max_column + 1):
                            cell = new_ws.cell(r, c)
                            val_str = str(cell.value or "").strip().upper()

                            # 匹配关键字
                            if "TOTAL DAP" in val_str:
                                # 1. 加粗关键字所在单元格
                                cell.font = Font(bold=True, name=cell.font.name, size=cell.font.size,
                                                 color=cell.font.color)

                                # 2. 从当前列 (c) 开始向右查找第一个有数值的单元格
                                for target_c in range(c + 1, new_ws.max_column + 1):
                                    target_cell = new_ws.cell(r, target_c)
                                    if target_cell.value is not None:
                                        # 找到第一个非空单元格，执行加粗
                                        target_cell.font = Font(bold=True, name=target_cell.font.name,
                                                                size=target_cell.font.size,
                                                                color=target_cell.font.color)
                                        print(f"   -> 已加粗第 {r} 行的数值单元格: {get_column_letter(target_c)}{r}")
                                        # 找到数值后通常这一行就处理完了，可以跳出
                                        break

                                found_dap = True
                                break
                        if found_dap: break
                except Exception as e:
                    print(f"   -> 加粗 TOTAL DAP 时出错: {e}")
                # =========================================================

        # 保存
        suffix = chr(64 + idx)
        if output_prefix.endswith('.xlsx'): output_prefix = output_prefix[:-5]
        parts = output_prefix.rsplit(' ', 1)
        out_path = f"{parts[0]}{suffix} {parts[1]}.xlsx" if len(parts) == 2 else f"{output_prefix}{suffix}.xlsx"
        new_wb.save(out_path)
        output_files.append(out_path)
        print(f"✅ {out_path} (样式修复完成)")

    return output_files