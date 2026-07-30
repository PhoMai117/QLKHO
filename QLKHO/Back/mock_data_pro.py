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

    cursor.execute("PRAGMA foreign_keys = ON;")

    # =======================================================
    # 1. TẠO CẤU TRÚC BẢNG 
    # =======================================================
    cursor.executescript('''
        CREATE TABLE USERS (
            User_ID VARCHAR(20) PRIMARY KEY,
            Ho_Ten NVARCHAR(100) NOT NULL,
            Mat_Khau VARCHAR(100) NOT NULL,
            Bo_Phan NVARCHAR(50) NOT NULL,
            Chuc_Vu NVARCHAR(50) NOT NULL,
            Trang_thai NVARCHAR(50) DEFAULT 'Hoạt động'
        );

        CREATE TABLE CUSTOMERS (
            Ma_KH VARCHAR(50) PRIMARY KEY, 
            Ten_KH NVARCHAR(255)
        );

        -- ĐÃ BỔ SUNG CỘT NCC CHO BẢNG PRODUCTS
        CREATE TABLE PRODUCTS (
            MSP VARCHAR(50) PRIMARY KEY, 
            Ten_SP NVARCHAR(255), 
            Quy_cach_Tui INT, 
            Quy_cach_Thung INT, 
            Don_Vi VARCHAR(50),
            NCC NVARCHAR(100) 
        );

        CREATE TABLE ORDERS (
            Ma_DH VARCHAR(50) PRIMARY KEY, 
            Ma_KH VARCHAR(50), 
            User_ID VARCHAR(20) NOT NULL,
            Ngay_tao DATETIME, 
            Ngay_du_kien_giao DATE, 
            Ngay_giao_thuc_te DATETIME, 
            Trang_thai_DH NVARCHAR(50),
            CONSTRAINT fk_orders_customers FOREIGN KEY (Ma_KH) REFERENCES CUSTOMERS(Ma_KH) ON UPDATE CASCADE,
            CONSTRAINT fk_orders_users FOREIGN KEY (User_ID) REFERENCES USERS(User_ID) ON UPDATE CASCADE
        );

        CREATE TABLE ORDER_DETAILS (
            Ma_PO VARCHAR(50) PRIMARY KEY, 
            Ma_DH VARCHAR(50), 
            MSP VARCHAR(50), 
            So_luong_Pcs INT, 
            CONSTRAINT fk_orderdetails_orders FOREIGN KEY (Ma_DH) REFERENCES ORDERS(Ma_DH) ON DELETE CASCADE,
            CONSTRAINT fk_orderdetails_products FOREIGN KEY (MSP) REFERENCES PRODUCTS(MSP) ON UPDATE CASCADE
        );

        CREATE TABLE BOXES (
            Box_ID VARCHAR(50) PRIMARY KEY, 
            MSP VARCHAR(50), 
            Ngay_dong_goi DATETIME, 
            Ngay_nhap_kho DATETIME, 
            So_luong_Tui_Thuc_Te INT, 
            Trang_thai NVARCHAR(50),
            Detail_ID VARCHAR(50) DEFAULT NULL,     -- Cột chứa mã PO khi xuất kho
            CONSTRAINT fk_boxes_products FOREIGN KEY (MSP) REFERENCES PRODUCTS(MSP) ON UPDATE CASCADE,
            CONSTRAINT fk_boxes_po FOREIGN KEY (Detail_ID) REFERENCES ORDER_DETAILS(Detail_ID) ON DELETE SET NULL
        );

        CREATE TABLE BAGS (
            Bag_ID VARCHAR(50) PRIMARY KEY, 
            Box_ID VARCHAR(50), 
            MSP VARCHAR(50), 
            NSX DATE, 
            So_Lot VARCHAR(50), 
            Trang_thai NVARCHAR(50),
            CONSTRAINT fk_bags_boxes FOREIGN KEY (Box_ID) REFERENCES BOXES(Box_ID) ON DELETE SET NULL,
            CONSTRAINT fk_bags_products FOREIGN KEY (MSP) REFERENCES PRODUCTS(MSP) ON UPDATE CASCADE
        );

        CREATE TABLE INVENTORY_TRANSACTIONS (
            Trans_ID VARCHAR(50) PRIMARY KEY, 
            Box_ID VARCHAR(50), 
            User_ID VARCHAR(20) NOT NULL,       
            Loai_giao_dich NVARCHAR(50), 
            Ma_Tham_Chieu VARCHAR(50), 
            Thoi_gian DATETIME, 
            Ma_Chung_Tu VARCHAR(50),
            CONSTRAINT fk_transactions_boxes FOREIGN KEY (Box_ID) REFERENCES BOXES(Box_ID) ON UPDATE CASCADE,
            CONSTRAINT fk_transactions_users FOREIGN KEY (User_ID) REFERENCES USERS(User_ID) ON UPDATE CASCADE
        );
    ''')

    # =======================================================
    # 2. KHỞI TẠO DỮ LIỆU CƠ BẢN
    # =======================================================
    users_data = [
        ('NTV-0001', 'Nguyễn Văn Trường', '1', 'Sản Xuất', 'Quản lý', 'Hoạt động'),
        ('NTV-0002', 'Trần Thị Lan', '1', 'Sản Xuất', 'Công nhân QC', 'Hoạt động'),
        ('NTV-0006', 'Lê Văn Khoa', '1', 'Kho', 'Thủ kho', 'Hoạt động')
    ]
    cursor.executemany("INSERT INTO USERS VALUES (?, ?, ?, ?, ?, ?)", users_data)

    # ĐÃ CẬP NHẬT DỮ LIỆU NCC (SONY, HUWAI, SAMSUNG)
    products = [
        ('SP-001', 'Mạch điều khiển A1', 125, 4, 'pcs', 'SONY'), 
        ('SP-002', 'Mạch điều khiển A2', 125, 4, 'pcs', 'SONY'),
        ('SP-003', 'Cảm biến nhiệt độ', 125, 4, 'pcs', 'SONY'), 
        ('SP-004', 'Cảm biến độ ẩm', 125, 4, 'pcs', 'SONY'),
        ('SP-005', 'Module Wifi', 125, 4, 'pcs', 'HUWAI'), 
        ('SP-006', 'Module Bluetooth', 125, 4, 'pcs', 'HUWAI'),
        ('SP-007', 'Vỏ nhựa ABS', 200, 2, 'pcs', 'HUWAI'), 
        ('SP-008', 'Vỏ nhôm tản nhiệt', 200, 2, 'pcs', 'SAMSUNG'),
        ('SP-009', 'Ốc vít mạ kẽm', 100, 5, 'pcs', 'SAMSUNG'), 
        ('SP-010', 'Dây cáp bọc dù', 100, 5, 'pcs', 'SAMSUNG')
    ]
    cursor.executemany("INSERT INTO PRODUCTS VALUES (?, ?, ?, ?, ?, ?)", products)
    
    customers = [(f'KH{str(i).zfill(2)}', f'Công ty Đối tác số {i}') for i in range(1, 7)]
    cursor.executemany("INSERT INTO CUSTOMERS VALUES (?, ?)", customers)

    print("Đang chạy mô phỏng nhập kho và xuất kho theo chuẩn PO mới...")

    start_date = datetime(2026, 1, 1)
    end_date = datetime(2026, 6, 30)
    current_date = start_date

    current_day_str = ""
    current_month_str = ""
    bag_seq = 1
    box_seq = 1
    order_seq = 1

    leftover_bags = {p[0]: [] for p in products}

    # =======================================================
    # 3. MÔ PHỎNG SẢN XUẤT VÀ TẠO ĐƠN THEO NGÀY
    # =======================================================
    while current_date <= end_date:
        date_str = current_date.strftime('%y%m%d')
        month_str = current_date.strftime('%y%m')

        if date_str != current_day_str:
            bag_seq = 1
            box_seq = 1
            current_day_str = date_str
        if month_str != current_month_str:
            order_seq = 1
            current_month_str = month_str
        
        # --- A. NHẬP KHO HÀNG NGÀY ---
        daily_products = random.sample(products, random.randint(2, 3))
        for prod in daily_products:
            msp = prod[0]
            qty_box = prod[3]
            lot_no = f"LOT-{date_str}01"
            total_bags_produced = random.randint(15, 25) 
            
            for _ in range(total_bags_produced):
                bag_id = f"B-{date_str}{str(bag_seq).zfill(3)}"
                bag_seq += 1
                cursor.execute("INSERT INTO BAGS VALUES (?, NULL, ?, ?, ?, 'Bình thường')", 
                               (bag_id, msp, current_date.strftime('%Y-%m-%d'), lot_no))
                leftover_bags[msp].append(bag_id)

            while len(leftover_bags[msp]) >= qty_box:
                bags_to_pack = leftover_bags[msp][:qty_box]
                leftover_bags[msp] = leftover_bags[msp][qty_box:] 
                
                box_id = f"T-{date_str}{str(box_seq).zfill(2)}"
                box_seq += 1
                time_pack = current_date + timedelta(hours=random.randint(8, 16))
                
                cursor.execute("INSERT INTO BOXES (Box_ID, MSP, Ngay_dong_goi, Ngay_nhap_kho, So_luong_Tui_Thuc_Te, Trang_thai) VALUES (?, ?, ?, ?, ?, 'Tồn kho')", 
                               (box_id, msp, time_pack.strftime('%Y-%m-%d %H:%M:%S'), time_pack.strftime('%Y-%m-%d %H:%M:%S'), qty_box))
                
                for b_id in bags_to_pack:
                    cursor.execute("UPDATE BAGS SET Box_ID=? WHERE Bag_ID=?", (box_id, b_id))
                
                trans_id = f"IN-{box_id}"
                cursor.execute("INSERT INTO INVENTORY_TRANSACTIONS VALUES (?, ?, 'NTV-0006', 'Nhập kho', NULL, ?, NULL)", 
                               (trans_id, box_id, time_pack.strftime('%Y-%m-%d %H:%M:%S')))
        
        # --- B. LÊN ĐƠN HÀNG & XUẤT KHO ---
        if random.random() < (3 / 30):
            order_id = f"DH-KH-{month_str}-{str(order_seq).zfill(2)}"
            order_seq += 1
            ma_kh = random.choice(customers)[0]
            creator_user = 'NTV-0006' 
            
            order_time = current_date + timedelta(hours=random.randint(8, 11))
            expected_date = current_date + timedelta(days=random.randint(2, 5))
            
            cursor.execute("INSERT INTO ORDERS VALUES (?, ?, ?, ?, ?, NULL, 'Chờ xuất kho')", 
                           (order_id, ma_kh, creator_user, order_time.strftime('%Y-%m-%d %H:%M:%S'), expected_date.strftime('%Y-%m-%d')))
            
            order_items = random.sample(products, random.randint(2, 4))
            order_items_details = []
            
            for idx, item in enumerate(order_items):
                msp = item[0]
                pcs_per_bag = item[2]
                pcs_per_box = pcs_per_bag * item[3]
                
                req_pcs = (random.randint(5, 15) * pcs_per_box) + random.choice([0, pcs_per_bag, pcs_per_bag*2])
                
                # --- ĐÃ SỬA: Chèn thêm (order_seq-1) vào để mã PO không bao giờ trùng ---
                po_id = f"PO-XYZ-{month_str}-{str(order_seq-1).zfill(2)}-{str(idx+1).zfill(2)}"
                
                cursor.execute("INSERT INTO ORDER_DETAILS VALUES (?, ?, ?, ?)", (po_id, order_id, msp, req_pcs))
                order_items_details.append((po_id, msp, req_pcs))
            
            can_fulfill = True
            for po_id, msp, req_pcs in order_items_details:
                cursor.execute('''SELECT SUM(b.So_luong_Tui_Thuc_Te * p.Quy_cach_Tui) 
                                  FROM BOXES b JOIN PRODUCTS p ON b.MSP = p.MSP 
                                  WHERE b.MSP=? AND b.Trang_thai='Tồn kho' AND b.Ngay_nhap_kho <= ?''', 
                               (msp, expected_date.strftime('%Y-%m-%d %H:%M:%S')))
                stock_avail = cursor.fetchone()[0] or 0
                if stock_avail < req_pcs:
                    can_fulfill = False
                    break
            
            if can_fulfill:
                for po_id, msp, req_pcs in order_items_details:
                    cursor.execute('''SELECT b.Box_ID, (b.So_luong_Tui_Thuc_Te * p.Quy_cach_Tui) as Pcs
                                      FROM BOXES b JOIN PRODUCTS p ON b.MSP = p.MSP 
                                      WHERE b.MSP=? AND b.Trang_thai='Tồn kho' AND b.Ngay_nhap_kho <= ? 
                                      ORDER BY b.Ngay_nhap_kho ASC''', 
                                   (msp, expected_date.strftime('%Y-%m-%d %H:%M:%S')))
                    boxes_to_export = cursor.fetchall()
                    accumulated = 0
                    
                    for b in boxes_to_export:
                        box_id = b[0]
                        pcs_in_box = b[1]
                        
                        cursor.execute("UPDATE BOXES SET Trang_thai='Đã xuất', Ma_PO=? WHERE Box_ID=?", (po_id, box_id))
                        
                        trans_out_id = f"OUT-{box_id}-{order_id}"
                        cursor.execute("INSERT INTO INVENTORY_TRANSACTIONS VALUES (?, ?, 'NTV-0006', 'Xuất kho', ?, ?, NULL)", 
                                       (trans_out_id, box_id, order_id, expected_date.strftime('%Y-%m-%d %H:%M:%S')))
                        
                        accumulated += pcs_in_box
                        if accumulated >= req_pcs:
                            break 
                
                cursor.execute("UPDATE ORDERS SET Trang_thai_DH='Đã giao', Ngay_giao_thuc_te=? WHERE Ma_DH=?", 
                               (expected_date.strftime('%Y-%m-%d %H:%M:%S'), order_id))

        current_date += timedelta(days=1)

    conn.commit()
    conn.close()
    print("HOÀN TẤT! Đã tạo xong Database tích hợp luồng chuẩn PO (Đã sửa cột NCC).")

if __name__ == '__main__':
    create_enterprise_mock_database()