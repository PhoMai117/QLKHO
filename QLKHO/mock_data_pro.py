import sqlite3
import random
from datetime import datetime, timedelta
import os

def create_enterprise_mock_database():
    db_path = 'data.db'
    if os.path.exists(db_path):
        try:
            os.remove(db_path)
            print("Đã dọn dẹp data.db cũ.")
        except PermissionError:
            print("LỖI: Hãy tắt Flask Server và DBeaver trước khi chạy!")
            return

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # 1. TẠO CẤU TRÚC BẢNG (Chuẩn hóa)
    cursor.executescript('''
        CREATE TABLE PRODUCTS (MSP VARCHAR(50) PRIMARY KEY, Ten_SP NVARCHAR(255), Quy_cach_Tui INT, Quy_cach_Thung INT, Don_Vi VARCHAR(50));
        CREATE TABLE CUSTOMERS (Ma_KH VARCHAR(50) PRIMARY KEY, Ten_KH NVARCHAR(255));
        CREATE TABLE BOXES (Box_ID VARCHAR(50) PRIMARY KEY, MSP VARCHAR(50), Ngay_dong_goi DATETIME, Ngay_nhap_kho DATETIME, So_luong_Tui_Thuc_Te INT, Trang_thai NVARCHAR(50));
        CREATE TABLE BAGS (Bag_ID VARCHAR(50) PRIMARY KEY, Box_ID VARCHAR(50), MSP VARCHAR(50), NSX DATE, So_Lot VARCHAR(50), Trang_thai NVARCHAR(50));
        CREATE TABLE ORDERS (Ma_DH VARCHAR(50) PRIMARY KEY, Ma_KH VARCHAR(50), Ngay_tao DATETIME, Ngay_giao DATETIME, Trang_thai_DH NVARCHAR(50));
        CREATE TABLE ORDER_DETAILS (Detail_ID VARCHAR(50) PRIMARY KEY, Ma_DH VARCHAR(50), MSP VARCHAR(50), So_luong_Thung INT);
        CREATE TABLE INVENTORY_TRANSACTIONS (Trans_ID VARCHAR(50) PRIMARY KEY, Box_ID VARCHAR(50), Loai_giao_dich NVARCHAR(50), Ma_Tham_Chieu VARCHAR(50), Thoi_gian DATETIME, Ma_Chung_Tu VARCHAR(50));
    ''')

    # KHỞI TẠO SẢN PHẨM & KHÁCH HÀNG
    products = [
        ('SP-001', 'Mạch điều khiển A1', 125, 4, 'pcs'), ('SP-002', 'Mạch điều khiển A2', 125, 4, 'pcs'),
        ('SP-003', 'Cảm biến nhiệt độ', 125, 4, 'pcs'), ('SP-004', 'Cảm biến độ ẩm', 125, 4, 'pcs'),
        ('SP-005', 'Module Wifi', 125, 4, 'pcs'), ('SP-006', 'Module Bluetooth', 125, 4, 'pcs'),
        ('SP-007', 'Vỏ nhựa ABS', 200, 2, 'pcs'), ('SP-008', 'Vỏ nhôm tản nhiệt', 200, 2, 'pcs'),
        ('SP-009', 'Ốc vít mạ kẽm', 100, 5, 'pcs'), ('SP-010', 'Dây cáp bọc dù', 100, 5, 'pcs')
    ]
    cursor.executemany("INSERT INTO PRODUCTS VALUES (?, ?, ?, ?, ?)", products)
    customers = [(f'KH{str(i).zfill(2)}', f'Công ty Đối tác số {i}') for i in range(1, 7)]
    cursor.executemany("INSERT INTO CUSTOMERS VALUES (?, ?)", customers)

    print("Đang chạy cỗ máy thời gian giả lập ")

    start_date = datetime(2026, 1, 1)
    end_date = datetime(2026, 6, 30)
    current_date = start_date

    # Biến kiểm soát reset theo thời gian
    current_day_str = ""
    current_month_str = ""
    bag_seq = 1
    box_seq = 1
    order_seq = 1

    # Kho chứa "Túi tồn cuối ngày" chưa đóng thùng
    leftover_bags = {p[0]: [] for p in products}

    # CHẠY VÒNG LẶP THEO TỪNG NGÀY
    while current_date <= end_date:
        date_str = current_date.strftime('%y%m%d')
        month_str = current_date.strftime('%y%m')

        # Logic xoay vòng (Reset sequence)
        if date_str != current_day_str:
            bag_seq = 1
            box_seq = 1
            current_day_str = date_str
        if month_str != current_month_str:
            order_seq = 1
            current_month_str = month_str
        
        # --- A. SẢN XUẤT HÀNG NGÀY ---
        daily_products = random.sample(products, random.randint(2, 3))
        
        for prod in daily_products:
            msp, _, _, qty_box, _ = prod
            num_lots = random.randint(1, 2)
            
            for lot_idx in range(num_lots):
                lot_no = f"LOT-{date_str}{str(lot_idx+1).zfill(2)}"
                # Sản xuất ra ngẫu nhiên số Túi (cố tình để số lẻ không chia hết cho quy cách Thùng)
                total_bags_produced = random.randint(15, 25) 
                
                # In tem Túi trước (chưa có Box_ID)
                for _ in range(total_bags_produced):
                    bag_id = f"B-{date_str}{str(bag_seq).zfill(3)}"
                    bag_seq += 1
                    # Lưu tạm túi vào DB với Box_ID = NULL
                    cursor.execute("INSERT INTO BAGS VALUES (?, NULL, ?, ?, ?, 'Bình thường')", 
                                   (bag_id, msp, current_date.strftime('%Y-%m-%d'), lot_no))
                    # Đưa vào kho chờ đóng gói
                    leftover_bags[msp].append(bag_id)

            # TIẾN HÀNH ĐÓNG THÙNG (Lấy từ kho chờ đóng gói)
            # Đây là lúc xảy ra "Thùng Mix": Lấy túi tồn của ngày hôm qua ghép với túi hôm nay
            while len(leftover_bags[msp]) >= qty_box:
                # Bốc đúng số lượng túi bằng quy cách Thùng
                bags_to_pack = leftover_bags[msp][:qty_box]
                leftover_bags[msp] = leftover_bags[msp][qty_box:] # Xóa túi đã bốc khỏi hàng chờ
                
                # Sinh mã Thùng
                box_id = f"T-{date_str}{str(box_seq).zfill(2)}"
                box_seq += 1
                time_pack = current_date + timedelta(hours=random.randint(8, 16))
                
                # Cập nhật Thùng vào DB
                cursor.execute("INSERT INTO BOXES VALUES (?, ?, ?, ?, ?, 'Tồn kho')", 
                               (box_id, msp, time_pack.strftime('%Y-%m-%d %H:%M:%S'), time_pack.strftime('%Y-%m-%d %H:%M:%S'), qty_box))
                
                # Gắn Box_ID cho các Túi con (Update quan hệ Cha-Con)
                for b_id in bags_to_pack:
                    cursor.execute("UPDATE BAGS SET Box_ID=? WHERE Bag_ID=?", (box_id, b_id))
                
                # Ghi nhận Nhập kho
                trans_id = f"IN-{box_id}"
                cursor.execute("INSERT INTO INVENTORY_TRANSACTIONS VALUES (?, ?, 'Nhập kho', NULL, ?, NULL)", 
                               (trans_id, box_id, time_pack.strftime('%Y-%m-%d %H:%M:%S')))
        
        # --- B. XỬ LÝ ĐƠN HÀNG (Reset theo tháng) ---
        if random.random() < (3 / 30):
            order_id = f"DH-{month_str}{str(order_seq).zfill(2)}"
            order_seq += 1
            ma_kh = random.choice(customers)[0]
            order_time = current_date + timedelta(hours=random.randint(8, 11))
            
            order_items = random.sample(products, random.randint(3, 5))
            cursor.execute("INSERT INTO ORDERS VALUES (?, ?, ?, NULL, 'Chờ xử lý')", 
                           (order_id, ma_kh, order_time.strftime('%Y-%m-%d %H:%M:%S')))
            
            for idx, item in enumerate(order_items):
                msp = item[0]
                req_qty = random.randint(10, 20)
                po_id = f"PO-{month_str}{str(order_seq-1).zfill(2)}-{str(idx+1).zfill(2)}"
                cursor.execute("INSERT INTO ORDER_DETAILS VALUES (?, ?, ?, ?)", (po_id, order_id, msp, req_qty))
                
            # LOGIC XUẤT KHO FIFO 
            export_time = current_date + timedelta(days=random.randint(2, 5))
            can_fulfill = True
            for item in order_items:
                msp = item[0]
                cursor.execute("SELECT So_luong_Thung FROM ORDER_DETAILS WHERE Ma_DH=? AND MSP=?", (order_id, msp))
                req_qty = cursor.fetchone()[0]
                cursor.execute("SELECT COUNT(*) FROM BOXES WHERE MSP=? AND Trang_thai='Tồn kho' AND Ngay_nhap_kho <= ?", 
                               (msp, export_time.strftime('%Y-%m-%d %H:%M:%S')))
                stock_avail = cursor.fetchone()[0]
                if stock_avail < req_qty:
                    can_fulfill = False
                    break
            
            if can_fulfill:
                for item in order_items:
                    msp = item[0]
                    cursor.execute("SELECT So_luong_Thung FROM ORDER_DETAILS WHERE Ma_DH=? AND MSP=?", (order_id, msp))
                    req_qty = cursor.fetchone()[0]
                    cursor.execute("SELECT Box_ID FROM BOXES WHERE MSP=? AND Trang_thai='Tồn kho' AND Ngay_nhap_kho <= ? ORDER BY Ngay_nhap_kho ASC LIMIT ?", 
                                   (msp, export_time.strftime('%Y-%m-%d %H:%M:%S'), req_qty))
                    boxes_to_export = cursor.fetchall()
                    for b in boxes_to_export:
                        box_id = b[0]
                        cursor.execute("UPDATE BOXES SET Trang_thai='Đã xuất' WHERE Box_ID=?", (box_id,))
                        trans_out_id = f"OUT-{box_id}-{order_id}"
                        cursor.execute("INSERT INTO INVENTORY_TRANSACTIONS VALUES (?, ?, 'Xuất kho', ?, ?, NULL)", 
                                       (trans_out_id, box_id, order_id, export_time.strftime('%Y-%m-%d %H:%M:%S')))
                
                cursor.execute("UPDATE ORDERS SET Trang_thai_DH='Đã giao', Ngay_giao=? WHERE Ma_DH=?", 
                               (export_time.strftime('%Y-%m-%d %H:%M:%S'), order_id))

        current_date += timedelta(days=1)

    conn.commit()
    conn.close()
    print("HOÀN TẤT! Đã sinh dữ liệu hoàn hảo với ID reset xoay vòng và luồng Thùng Mix các Lot/Ngày.")

if __name__ == '__main__':
    create_enterprise_mock_database()