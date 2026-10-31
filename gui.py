import tkinter as tk
from tkinter import ttk, messagebox
from threading import Thread
from logic import chon_file_excel, lay_danh_sach_sheet, xu_ly_tat_ca_ke_hoach

class LoginWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Đăng nhập hệ thống - XLKH")
        self.root.geometry("350x220")
        self.root.resizable(False, False)
        
        self.center_window(350, 220)
        
        lbl_title = tk.Label(root, text="ĐĂNG NHẬP HỆ THỐNG", font=("Arial", 14, "bold"), fg="#2c3e50")
        lbl_title.pack(pady=15)
        
        frame_form = tk.Frame(root)
        frame_form.pack(pady=5)
        
        tk.Label(frame_form, text="Tài khoản:", font=("Arial", 10)).grid(row=0, column=0, sticky="w", pady=5)
        self.entry_user = tk.Entry(frame_form, font=("Arial", 10), width=20)
        self.entry_user.grid(row=0, column=1, pady=5, padx=5)
        
        tk.Label(frame_form, text="Mật khẩu:", font=("Arial", 10)).grid(row=1, column=0, sticky="w", pady=5)
        self.entry_pass = tk.Entry(frame_form, font=("Arial", 10), width=20, show="*")
        self.entry_pass.grid(row=1, column=1, pady=5, padx=5)
        
        btn_login = tk.Button(root, text="Đăng nhập", font=("Arial", 10, "bold"), bg="#27ae60", fg="white", width=15, command=self.check_login)
        btn_login.pack(pady=15)

    def center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width/2) - (width/2)
        y = (screen_height/2) - (height/2)
        self.root.geometry('%dx%d+%d+%d' % (width, height, x, y))

    def check_login(self):
        username = self.entry_user.get().strip()
        password = self.entry_pass.get().strip()
        
        # Cho phép đăng nhập linh hoạt hoặc theo tài khoản của bạn
        if username and password: 
            self.root.destroy()
            main_root = tk.Tk()
            AppGUI(main_root)
            main_root.mainloop()
        else:
            messagebox.showerror("Lỗi đăng nhập", "Vui lòng nhập đầy đủ tài khoản và mật khẩu!")


class AppGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Phần mềm Kế hoạch Sản xuất")
        self.root.geometry("540x500")
        self.root.resizable(False, False)
        self.file_path = ""
        
        self.center_window(540, 500)
        
        # Tiêu đề
        lbl_header = tk.Label(root, text="QUẢN LÝ KẾ HOẠCH SẢN XUẤT", font=("Arial", 14, "bold"), fg="#2980b9")
        lbl_header.pack(pady=15)
        
        # Khung chọn file
        frame_file = tk.Frame(root)
        frame_file.pack(pady=5, fill="x", padx=20)
        
        self.btn_chon = tk.Button(frame_file, text="Chọn File Excel", command=self.chon_file, bg="#3498db", fg="white", font=("Arial", 9, "bold"))
        self.btn_chon.pack(side="left")
        
        self.lbl_path = tk.Label(frame_file, text="Chưa chọn file nào", fg="gray", anchor="w")
        self.lbl_path.pack(side="left", padx=10, fill="x", expand=True)
        
        # Khung cấu hình Sheet và Tháng
        frame_config = tk.Frame(root)
        frame_config.pack(pady=10, fill="x", padx=20)
        
        tk.Label(frame_config, text="Chọn Sheet nguồn:", font=("Arial", 10)).grid(row=0, column=0, sticky="w", pady=5)
        self.sheet_combobox = ttk.Combobox(frame_config, state="readonly", width=25)
        self.sheet_combobox.grid(row=0, column=1, pady=5, padx=10, sticky="w")
        
        tk.Label(frame_config, text="Chọn Tháng:", font=("Arial", 10)).grid(row=1, column=0, sticky="w", pady=5)
        self.month_combobox = ttk.Combobox(frame_config, state="readonly", width=10, values=[str(i) for i in range(1, 13)])
        self.month_combobox.grid(row=1, column=1, pady=5, padx=10, sticky="w")
        self.month_combobox.current(9)  # Mặc định tháng 10
        
        # KHUNG TÙY CHỌN XỬ LÝ (Ép vỉ, Pha chế, Lịch SP)
        frame_options = tk.LabelFrame(root, text=" Tùy chọn nội dung xuất kết quả ", font=("Arial", 10, "bold"), fg="#2c3e50")
        frame_options.pack(pady=10, fill="x", padx=20)
        
        self.var_ep_vi = tk.BooleanVar(value=True)
        self.var_pha_che = tk.BooleanVar(value=True)
        self.var_lich_sp = tk.BooleanVar(value=True)
        
        tk.Checkbutton(frame_options, text="Xuất sheet Ép vỉ", variable=self.var_ep_vi, font=("Arial", 9)).pack(anchor="w", padx=10, pady=2)
        tk.Checkbutton(frame_options, text="Xuất sheet Pha chế", variable=self.var_pha_che, font=("Arial", 9)).pack(anchor="w", padx=10, pady=2)
        tk.Checkbutton(frame_options, text="Xuất sheet Lịch SP (Tổng hợp)", variable=self.var_lich_sp, font=("Arial", 9)).pack(anchor="w", padx=10, pady=2)
        
        # Trạng thái và Thanh Loading
        self.lbl_status = tk.Label(root, text="Sẵn sàng", fg="gray", font=("Arial", 9, "italic"))
        self.lbl_status.pack(pady=5)
        
        self.progress = ttk.Progressbar(root, orient="horizontal", length=480, mode="determinate")
        self.progress.pack(pady=5)
        
        # Nút chạy xử lý
        self.btn_chay = tk.Button(root, text="Bắt đầu Xử lý & Xuất File", command=self.chay_ngam_tien_trinh, bg="#27ae60", fg="white", font=("Arial", 11, "bold"), state=tk.DISABLED)
        self.btn_chay.pack(pady=15)

    def center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width/2) - (width/2)
        y = (screen_height/2) - (height/2)
        self.root.geometry('%dx%d+%d+%d' % (width, height, x, y))

    def chon_file(self):
        path = chon_file_excel()
        if path:
            self.file_path = path
            filename = path.split("/")[-1].split("\\")[-1]
            self.lbl_path.config(text=filename, fg="#27ae60")
            
            sheets = lay_danh_sach_sheet(path)
            if sheets:
                self.sheet_combobox['values'] = sheets
                self.sheet_combobox.current(0)
                self.btn_chay.config(state=tk.NORMAL)
            else:
                messagebox.showwarning("Cảnh báo", "Không đọc được danh sách Sheet từ file này!")

    def chay_ngam_tien_trinh(self):
        # Kiểm tra xem người dùng có chọn ít nhất 1 mục nào không
        if not (self.var_ep_vi.get() or self.var_pha_che.get() or self.var_lich_sp.get()):
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ít nhất một nội dung cần xuất kết quả!")
            return

        self.btn_chay.config(state=tk.DISABLED)
        self.btn_chon.config(state=tk.DISABLED)
        self.progress['value'] = 0
        
        t = Thread(target=self.thuc_hien_xu_ly)
        t.daemon = True
        t.start()

    def thuc_hien_xu_ly(self):
        try:
            target_month = int(self.month_combobox.get())
            sheet_src = self.sheet_combobox.get()
            
            # Truyền thêm các cờ lựa chọn sang logic.py
            output_file = xu_ly_tat_ca_ke_hoach(
                self.file_path, 
                target_month, 
                sheet_src, 
                export_ep_vi=self.var_ep_vi.get(),
                export_pha_che=self.var_pha_che.get(),
                export_lich=self.var_lich_sp.get(),
                progress_callback=self.cap_nhat_giao_dien_loading
            )
            
            self.lbl_status.config(text="Hoàn thành xuất file!", fg="green")
            messagebox.showinfo("Thành công", f"Đã xuất file kết quả sạch sẽ tại:\n{output_file}")
            
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))
            self.lbl_status.config(text="Lỗi xử lý!", fg="red")
        finally:
            self.btn_chay.config(state=tk.NORMAL)
            self.btn_chon.config(state=tk.NORMAL)

    def cap_nhat_giao_dien_loading(self, percent, message):
        self.progress['value'] = percent
        self.lbl_status.config(text=message)
        self.root.update_idletasks()