import re
from datetime import datetime, timedelta
import openpyxl


def lay_danh_sach_sheet(file_path):
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True)
        names = wb.sheetnames
        wb.close()
        return names
    except Exception as e:
        print(f"Lỗi đọc sheet: {e}")
        return []

def lay_danh_sach_cot_theo_thang(ws_src, target_month):
    max_col = ws_src.max_column
    temp_dates = [] 
    current_year = datetime.now().year
    
    date_row = 5
    for r in range(1, 8):
        count_dates = 0
        for col in range(3, max_col + 1):
            val = ws_src.cell(row=r, column=col).value
            if isinstance(val, datetime) or (isinstance(val, (int, float)) and val > 30000):
                count_dates += 1
        if count_dates >= 3:
            date_row = r
            break

    for col in range(3, max_col + 1):
        val = ws_src.cell(row=date_row, column=col).value
        if val is not None:
            day_num = None
            if isinstance(val, datetime):
                if val.month == target_month:
                    day_num = val.day
            elif isinstance(val, (int, float)) and val > 30000:
                try:
                    base_date = datetime(1899, 12, 30)
                    dt_converted = base_date + timedelta(days=val)
                    if dt_converted.month == target_month:
                        day_num = dt_converted.day
                except Exception:
                    pass
            else:
                text_val = str(val).strip()
                match_day = re.search(r'\d+', text_val)
                if match_day:
                    d = int(match_day.group())
                    if 1 <= d <= 31 and d != target_month: 
                        day_num = d

            if day_num is not None:
                try:
                    dt_obj = datetime(current_year, target_month, day_num)
                    temp_dates.append((day_num, col, dt_obj))
                except ValueError:
                    continue

    if not temp_dates:
        return None, None

    temp_dates.sort(key=lambda x: x[0])
    date_columns = {} 
    formatted_dates_list = []   
    
    for day_num, col_idx, dt_obj in temp_dates:
        date_columns[col_idx] = dt_obj
        if dt_obj not in formatted_dates_list:
            formatted_dates_list.append(dt_obj)
            
    return date_columns, formatted_dates_list

def tim_dong_may_va_pha_che(ws_src):
    ep_vi_rows = {}
    pha_che_rows = []

    for r in range(1, ws_src.max_row + 1):
        val_a = str(ws_src.cell(row=r, column=1).value or "").lower()
        val_b = str(ws_src.cell(row=r, column=2).value or "").lower()
        combined = f"{val_a} {val_b}"

        if "ép vỉ" in combined or "ep vi" in combined:
            match_m = re.search(r'\d+', combined)
            m_id = match_m.group() if match_m else str(len(ep_vi_rows) + 1)
            ep_vi_rows[r] = m_id
        elif "pha chế" in combined or "pha che" in combined:
            pha_che_rows.append(r)

    if not ep_vi_rows:
        ep_vi_rows = {6: '1', 7: '2', 8: '3'}
    if not pha_che_rows:
        pha_che_rows = [9]

    return ep_vi_rows, pha_che_rows

def xu_ly_tat_ca_ke_hoach(file_path, target_month, sheet_src, 
                          export_ep_vi=True, export_pha_che=True, export_lich=True, 
                          progress_callback=None):
    
    if progress_callback:
        progress_callback(10, "Đang mở và đọc file Excel...")
        
    wb_src = openpyxl.load_workbook(file_path)
    if sheet_src not in wb_src.sheetnames:
        raise ValueError(f"Không tìm thấy sheet '{sheet_src}'!")
    ws_src = wb_src[sheet_src]

    if progress_callback:
        progress_callback(30, "Đang phân tích các cột ngày tháng...")
    
    date_columns, formatted_dates_list = lay_danh_sach_cot_theo_thang(ws_src, target_month)
    if not date_columns:
        raise ValueError(f"Không tìm thấy cột ngày nào thuộc Tháng {target_month}!")

    ep_vi_rows, pha_che_rows = tim_dong_may_va_pha_che(ws_src)
    pattern = re.compile(r'([^\(]+?)\s*\(\s*([\d\.]+)\s*([^\)]+)\s*\)')

    # --- 1. XỬ LÝ ÉP VỈ (Chỉ chạy khi được tick chọn) ---
    ep_vi_rows_data = []
    if export_ep_vi:
        if progress_callback:
            progress_callback(40, "Đang xử lý dữ liệu Ép vỉ...")
        
        ep_vi_data_dict = {}
        date_to_idx = {dt: i for i, dt in enumerate(formatted_dates_list)}
        for row_idx, machine_id in ep_vi_rows.items():
            for col_idx, dt_obj in date_columns.items():
                cell_value = ws_src.cell(row=row_idx, column=col_idx).value
                if not cell_value: continue
                matches = pattern.findall(str(cell_value).replace('\n', ' '))
                for raw_name, raw_qty, raw_unit in matches:
                    prod_name_clean = re.sub(r'\s+', ' ', raw_name.strip(' ,')).lower()
                    if not prod_name_clean: continue
                    qty_clean = int(raw_qty.replace('.', ''))
                    unit_clean = raw_unit.strip().lower()
                    if prod_name_clean not in ep_vi_data_dict:
                        ep_vi_data_dict[prod_name_clean] = {'unit': unit_clean, 'machines': set(), 'daily_qty': {}}
                    ep_vi_data_dict[prod_name_clean]['machines'].add(machine_id)
                    ep_vi_data_dict[prod_name_clean]['daily_qty'][dt_obj] = ep_vi_data_dict[prod_name_clean]['daily_qty'].get(dt_obj, 0) + qty_clean

        for prod_name, info in ep_vi_data_dict.items():
            daily_dict = info['daily_qty']
            active_dates = sorted(daily_dict.keys(), key=lambda d: date_to_idx[d])
            if not active_dates: continue
            active_indices = [date_to_idx[d] for d in active_dates]
            groups, current_group = [], [active_indices[0]]
            for idx in active_indices[1:]:
                if idx == current_group[-1] + 1: current_group.append(idx)
                else: groups.append(current_group); current_group = [idx]
            groups.append(current_group)
            new_daily_qty = {}
            for group in groups:
                batch_total = max([daily_dict[formatted_dates_list[idx]] for idx in group])
                qty_per_day = batch_total / len(group)
                for idx in group: new_daily_qty[formatted_dates_list[idx]] = round(qty_per_day)
            info['daily_qty'] = new_daily_qty

        for prod_name, info in ep_vi_data_dict.items():
            sorted_machines = sorted(list(info['machines']), key=lambda x: int(x) if x.isdigit() else x)
            row_data = {'Tên sản phẩm': prod_name, 'máy ép vỉ': '|'.join(sorted_machines), 'Đơn vị': info['unit']}
            total_month = 0
            for dt_obj in formatted_dates_list:
                qty = info['daily_qty'].get(dt_obj, 0)
                row_data[dt_obj] = qty if qty > 0 else ''
                total_month += qty
            row_data['Tổng'] = total_month
            ep_vi_rows_data.append(row_data)

    # --- 2. XỬ LÝ PHA CHẾ (Chỉ chạy khi được tick chọn) ---
    pha_che_rows_data = []
    if export_pha_che:
        if progress_callback:
            progress_callback(60, "Đang xử lý dữ liệu Pha chế...")
            
        pha_che_data_dict = {}
        for pha_che_row in pha_che_rows:
            for col_idx, dt_obj in date_columns.items():
                cell_value = ws_src.cell(row=pha_che_row, column=col_idx).value
                if not cell_value: continue
                matches = pattern.findall(str(cell_value).replace('\n', ' '))
                for raw_name, raw_qty, raw_unit in matches:
                    prod_name_clean = re.sub(r'\s+', ' ', raw_name.strip(' ,')).lower()
                    if not prod_name_clean: continue
                    qty_clean = int(raw_qty.replace('.', ''))
                    unit_clean = raw_unit.strip().lower() if raw_unit.strip() else None
                    if prod_name_clean not in pha_che_data_dict:
                        pha_che_data_dict[prod_name_clean] = {'unit': unit_clean, 'daily_qty': {}}
                    elif not pha_che_data_dict[prod_name_clean]['unit'] and unit_clean:
                        pha_che_data_dict[prod_name_clean]['unit'] = unit_clean
                    pha_che_data_dict[prod_name_clean]['daily_qty'][dt_obj] = pha_che_data_dict[prod_name_clean]['daily_qty'].get(dt_obj, 0) + qty_clean

        for prod_name, info in pha_che_data_dict.items():
            row_data = {'Tên sản phẩm': prod_name, 'Đơn vị': info['unit'] if info['unit'] is not None else ''}
            total_month = 0
            for dt_obj in formatted_dates_list:
                qty = info['daily_qty'].get(dt_obj, 0)
                row_data[dt_obj] = qty if qty > 0 else ''
                total_month += qty
            row_data['Tổng'] = total_month
            pha_che_rows_data.append(row_data)

    # --- 3. LỊCH TỔNG HỢP (Chỉ chạy khi được tick chọn) ---
    lich_rows_data = []
    if export_lich:
        if progress_callback:
            progress_callback(75, "Đang xử lý Lịch tổng hợp...")
            
        ep_vi_dates, pha_che_dates, thanh_pham_dates = {}, {}, {}
        for row_idx in ep_vi_rows.keys():
            for col_idx, dt_obj in date_columns.items():
                cell_value = ws_src.cell(row=row_idx, column=col_idx).value
                if not cell_value: continue
                for raw_name, _, _ in pattern.findall(str(cell_value).replace('\n', ' ')):
                    p_name = re.sub(r'\s+', ' ', raw_name.strip(' ,')).lower()
                    if p_name:
                        if p_name not in ep_vi_dates: ep_vi_dates[p_name] = set()
                        ep_vi_dates[p_name].add(dt_obj)

        for pha_che_row in pha_che_rows:
            for col_idx, dt_obj in date_columns.items():
                cell_value = ws_src.cell(row=pha_che_row, column=col_idx).value
                if not cell_value: continue
                for raw_name, _, _ in pattern.findall(str(cell_value).replace('\n', ' ')):
                    p_name = re.sub(r'\s+', ' ', raw_name.strip(' ,')).lower()
                    if p_name:
                        if p_name not in pha_che_dates: pha_che_dates[p_name] = set()
                        pha_che_dates[p_name].add(dt_obj)
                        tp_dt = dt_obj + timedelta(days=13)
                        if p_name not in thanh_pham_dates: thanh_pham_dates[p_name] = set()
                        thanh_pham_dates[p_name].add(tp_dt)

        all_products = sorted(list(set(list(ep_vi_dates.keys()) + list(pha_che_dates.keys()))))
        for p_name in all_products:
            d_ep_vi = sorted(list(ep_vi_dates.get(p_name, set())))
            d_pha_che = sorted(list(pha_che_dates.get(p_name, set())))
            d_thanh_pham = sorted(list(thanh_pham_dates.get(p_name, set())))
            lich_rows_data.append({
                'Tên sản phẩm': p_name,
                'Ngày ép vỉ': ", ".join([d.strftime('%d/%m/%Y') for d in d_ep_vi]) if d_ep_vi else "",
                'Ngày pha chế': ", ".join([d.strftime('%d/%m/%Y') for d in d_pha_che]) if d_pha_che else "",
                'Ngày dự kiến thành phẩm': ", ".join([d.strftime('%d/%m/%Y') for d in d_thanh_pham]) if d_thanh_pham else ""
            })

    wb_src.close()

    if progress_callback:
        progress_callback(85, "Đang khởi tạo file kết quả...")

    # --- KHỞI TẠO FILE KẾT QUẢ MỚI VÀ GHI CÁC SHEET ĐƯỢC CHỌN ---
    wb_dst = openpyxl.Workbook()
    default_sheet = wb_dst.active
    str_dates_list = [dt.strftime('%d/%m/%Y') for dt in formatted_dates_list]

    # Chỉ tạo sheet Ép vỉ nếu người dùng tích chọn
    if export_ep_vi:
        ws_ep = wb_dst.create_sheet(title=f"Chuyển đổi ép vỉ (Tháng {target_month})")
        ws_ep.append(['Tên sản phẩm', 'máy ép vỉ', 'Đơn vị'] + str_dates_list + ['Tổng'])
        for row in ep_vi_rows_data:
            row_vals = [row['Tên sản phẩm'], row['máy ép vỉ'], row['Đơn vị']]
            for dt_obj in formatted_dates_list: row_vals.append(row.get(dt_obj, ''))
            row_vals.append(row['Tổng'])
            ws_ep.append(row_vals)

    # Chỉ tạo sheet Pha chế nếu người dùng tích chọn
    if export_pha_che:
        ws_pc = wb_dst.create_sheet(title=f"Chuyển đổi pha chế (Tháng {target_month})")
        completion_str_list = [(dt + timedelta(days=13)).strftime('%d/%m/%Y') for dt in formatted_dates_list]
        ws_pc.append(['Ngày dự kiến thành phẩm', ''] + completion_str_list + [''])
        ws_pc.append(['Tên sản phẩm', 'Đơn vị'] + str_dates_list + ['Tổng'])
        for row in pha_che_rows_data:
            row_vals = [row['Tên sản phẩm'], row['Đơn vị']]
            for dt_obj in formatted_dates_list: row_vals.append(row.get(dt_obj, ''))
            row_vals.append(row['Tổng'])
            ws_pc.append(row_vals)

    # Chỉ tạo sheet Lịch SP nếu người dùng tích chọn
    if export_lich:
        ws_lich = wb_dst.create_sheet(title=f"Lịch SP (Tháng {target_month})")
        ws_lich.append(['Tên sản phẩm', 'Ngày ép vỉ', 'Ngày pha chế', 'Ngày dự kiến thành phẩm'])
        for item in lich_rows_data:
            ws_lich.append([item['Tên sản phẩm'], item['Ngày ép vỉ'], item['Ngày pha chế'], item['Ngày dự kiến thành phẩm']])

    # Gỡ bỏ sheet trắng mặc định ban đầu
    wb_dst.remove(default_sheet)

    if progress_callback:
        progress_callback(95, "Đang lưu file kết quả...")

    output_path = file_path.rsplit('.', 1)[0] + "_Ket_Qua.xlsx"
    wb_dst.save(output_path)
    
    if progress_callback:
        progress_callback(100, "Xử lý thành công!")
        
    return output_path

# Giữ lại các hàm cũ tương thích nếu giao diện có gọi riêng lẻ từng cái
def xu_ly_ep_vi(file_path, target_month, sheet_src):
    return xu_ly_tat_ca_ke_hoach(file_path, target_month, sheet_src)

def xu_ly_pha_che(file_path, target_month, sheet_src):
    return xu_ly_tat_ca_ke_hoach(file_path, target_month, sheet_src)

def tao_sheet_lich_tong_hop(file_path, target_month, sheet_src):
    return xu_ly_tat_ca_ke_hoach(file_path, target_month, sheet_src)