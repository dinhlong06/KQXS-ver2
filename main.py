import tkinter as tk
from tkinter import ttk, messagebox
import datetime
from datetime import datetime, timedelta
from process_data import fetch_data
from expo_excel import write_to_excel, PROVINCE_CONFIG, open_excel_file, close_excel_file

class XoSoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Kết Quả Xổ Số Miền Nam")
        self.root.geometry("900x750")
        
        self.results = None
        self.current_date = datetime.now()
        self.auto_update = False
        self.auto_update_id = None
        
        self.create_widgets()
        self.show_results(datetime.now())  # Mặc định hiển thị hôm nay
        
    def create_widgets(self):
        # Frame chứa nút điều khiển
        control_frame = ttk.Frame(self.root, padding="10")
        control_frame.pack(fill=tk.X)
        
        # Nút chọn ngày
        ttk.Button(control_frame, text="HÔM QUA", command=lambda: self.show_results(datetime.now() - timedelta(days=1), 1)).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="HÔM NAY", command=lambda: self.show_results(datetime.now(), 0)).pack(side=tk.LEFT, padx=5)
        
        # Nút làm mới
        ttk.Button(control_frame, text="LÀM MỚI", command=self.refresh_data).pack(side=tk.LEFT, padx=5)
        
        # Nút cập nhật tự động
        self.auto_update_btn = ttk.Button(control_frame, text="BẬT CẬP NHẬT TỰ ĐỘNG", command=self.toggle_auto_update)
        self.auto_update_btn.pack(side=tk.LEFT, padx=5)
        
        # Nút xuất Excel
        ttk.Button(control_frame, text="XUẤT EXCEL", command=self.export_to_excel).pack(side=tk.RIGHT, padx=5)
        
        # Hiển thị ngày
        self.date_label = ttk.Label(self.root, text="", font=('Arial', 14, 'bold'))
        self.date_label.pack(pady=10)
        
        # Frame hiển thị kết quả
        result_frame = ttk.Frame(self.root)
        result_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Tạo Treeview
        self.tree = ttk.Treeview(result_frame, columns=("Giai", "Tinh1", "Tinh2", "Tinh3"), show="headings", height=20)
        
        # Cấu hình cột
        self.tree.heading("Giai", text="GIẢI", anchor=tk.W)
        self.tree.heading("Tinh1", text="TỈNH 1", anchor=tk.CENTER)
        self.tree.heading("Tinh2", text="TỈNH 2", anchor=tk.CENTER)
        self.tree.heading("Tinh3", text="TỈNH 3", anchor=tk.CENTER)
        
        self.tree.column("Giai", width=120, anchor=tk.W)
        self.tree.column("Tinh1", width=180, anchor=tk.CENTER)
        self.tree.column("Tinh2", width=180, anchor=tk.CENTER)
        self.tree.column("Tinh3", width=180, anchor=tk.CENTER)
        
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Thanh cuộn
        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def show_results(self, date, day_ago=0):
        self.current_date = date
        weekday = date.weekday()

        date_str = date.strftime("%d-%m-%Y")
        # Cập nhật nhãn ngày
        self.date_label.config(text=f"KẾT QUẢ XỔ SỐ NGÀY {date_str}")
        
        # Lấy dữ liệu
        url = (f"https://www.minhngoc.net/ket-qua-xo-so/mien-nam/{date_str}.html" 
           if day_ago != 0 else "https://www.minhngoc.net/xo-so-truc-tiep/mien-nam.html")
        self.results = fetch_data(url, date)
        
        if not self.results:
            messagebox.showerror("Lỗi", f"Không tìm thấy kết quả cho ngày {date_str}")
            return
        
        # Xóa dữ liệu cũ
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Hiển thị kết quả
        province_names = PROVINCE_CONFIG["display_names"][weekday]
        self.tree.heading("Tinh1", text=province_names[0])
        self.tree.heading("Tinh2", text=province_names[1])
        self.tree.heading("Tinh3", text=province_names[2])
        
        prize_order = [
            ("Giải tám", "giai8"),
            ("Giải bảy", "giai7"),
            ("Giải sáu", "giai6"),
            ("Giải năm", "giai5"),
            ("Giải tư", "giai4"),
            ("Giải ba", "giai3"),
            ("Giải nhì", "giai2"),
            ("Giải nhất", "giai1"),
            ("Giải ĐB", "giaidb")
        ]
        
        for prize_name, prize_key in prize_order:
            # Thêm tên giải
            item = self.tree.insert("", tk.END, values=(prize_name, "", "", ""))
            self.tree.item(item, tags=('header',))
            
            # Thêm kết quả
            max_lines = max(len(self.results[i].get(prize_key, [])) for i in range(3))
            for line in range(max_lines):
                row_data = [""]
                for i in range(3):
                    numbers = self.results[i].get(prize_key, [])
                    row_data.append(numbers[line] if line < len(numbers) else "")
                
                item = self.tree.insert("", tk.END, values=row_data)
                if prize_key in ["giai8", "giaidb"]:
                    self.tree.item(item, tags=('bold',))
        
        # Định dạng
        self.tree.tag_configure('header', font=('Arial', 10, 'bold'))
        self.tree.tag_configure('bold', font=('Arial', 10, 'bold'))
        
    def refresh_data(self):
        """Làm mới dữ liệu hiện tại"""
        self.results = None
        self.show_results(self.current_date, 0)
        
        if (self.results and len(self.results) == 3 and all(len(result.get('giaidb', [])) > 0 for result in self.results)):
            if self.auto_update:
                self.toggle_auto_update()
            messagebox.showinfo("Thông báo", "Đã có đầy đủ kết quả Giải Đặc Biệt!")
    
    def toggle_auto_update(self):
        """Bật/tắt chế độ cập nhật tự động"""
        self.auto_update = not self.auto_update
        
        if self.auto_update:
            self.auto_update_btn.config(text="TẮT CẬP NHẬT TỰ ĐỘNG")
            self.start_auto_update()
        else:
            self.auto_update_btn.config(text="BẬT CẬP NHẬT TỰ ĐỘNG")
            self.stop_auto_update()
    
    def start_auto_update(self):
        """Bắt đầu cập nhật tự động mỗi 5 giây"""

        if self.auto_update:
            self.refresh_data()
            self.auto_update_id = self.root.after(5000, self.start_auto_update)  # 5 giây
    
    def stop_auto_update(self):
        """Dừng cập nhật tự động"""
        if self.auto_update_id:
            self.root.after_cancel(self.auto_update_id)
            self.auto_update_id = None
    
    def export_to_excel(self):
        if not self.results:
            messagebox.showwarning("Cảnh báo", "Không có dữ liệu để xuất Excel")
            return

        # Kiểm tra dữ liệu hợp lệ
        has_valid_data = any(
            any(len(tinh.get(prize_key, [])) > 0 
                for tinh in self.results 
                for prize_key in ["giaidb", "giai1", "giai2", "giai3", "giai4", "giai5", "giai6", "giai7", "giai8"]
            ) for tinh in self.results)
        
        if not has_valid_data:
            messagebox.showerror("Lỗi", "Không có dữ liệu kết quả xổ số hợp lệ để xuất Excel")
            return
        
        try:
            # Tự động đóng file nếu đang mở
            close_excel_file("example.xlsx")
            
            # Chờ file đóng hoàn toàn
            import time
            time.sleep(0.5)
            
            # Xuất dữ liệu vào Excel
            excel_date = self.current_date.strftime("%d/%m/%Y")
            write_to_excel(self.results, self.current_date.weekday(), excel_date)
            
            # Thông báo thành công
            messagebox.showinfo("Thành công", 
                f"Đã xuất kết quả ngày {excel_date} ra file Excel.\nFile sẽ tự động mở.")
            
            # Mở file Excel
            open_excel_file("example.xlsx")
            
        except Exception as e:
            messagebox.showerror("Lỗi", 
                f"Không thể xuất Excel: {str(e)}\nVui lòng thử lại sau.")

if __name__ == "__main__":
    root = tk.Tk()
    app = XoSoApp(root)
    root.mainloop()