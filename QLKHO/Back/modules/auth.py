import sqlite3
import os
from flask import Blueprint, request, jsonify

# Khởi tạo Blueprint cho Module Auth
auth_bp = Blueprint('auth', __name__)

# Hàm kết nối Database
def get_db_connection():
    db_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'data.db'))
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row  # Trả về kết quả dạng Dictionary/Row
    return conn

@auth_bp.route('/login', methods=['POST'])
def login():
    data = request.get_json()
    if not data:
        return jsonify({'success': False, 'message': 'Dữ liệu không hợp lệ!'}), 400

    user_id = data.get('user_id', '').strip()
    password = data.get('password', '').strip()

    if not user_id or not password:
        return jsonify({'success': False, 'message': 'Vui lòng nhập Mã NV và Mật khẩu!'}), 400

    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        # Kiểm tra thông tin tài khoản
        cursor.execute('''
            SELECT User_ID, Ho_Ten, Mat_Khau, Bo_Phan, Chuc_Vu, Trang_Thai 
            FROM USERS 
            WHERE User_ID = ?
        ''', (user_id,))
        
        user = cursor.fetchone()
        conn.close()

        if not user:
            return jsonify({'success': False, 'message': 'Mã nhân viên không tồn tại!'}), 401

        # Kiểm tra trạng thái tài khoản
        if user['Trang_Thai'] != 'Hoạt động':
            return jsonify({'success': False, 'message': 'Tài khoản của bạn đã bị khóa/vô hiệu hóa!'}), 403

        # Kiểm tra mật khẩu
        if user['Mat_Khau'] != password:
            return jsonify({'success': False, 'message': 'Mật khẩu không chính xác!'}), 401

        # Đăng nhập thành công -> Định hướng trang dựa theo Bộ phận
        redirect_page = "dashboard.html"
        if user['Bo_Phan'] == 'Sản Xuất':
            redirect_page = "production.html"
        elif user['Bo_Phan'] == 'Kho':
            redirect_page = "warehouse.html"
        elif user['Bo_Phan'] == 'Kinh Doanh':
            redirect_page = "sales.html"

        return jsonify({
            'success': True,
            'message': 'Đăng nhập thành công!',
            'user': {
                'user_id': user['User_ID'],
                'ho_ten': user['Ho_Ten'],
                'bo_phan': user['Bo_Phan'],
                'chuc_vu': user['Chuc_Vu']
            },
            'redirect': redirect_page
        }), 200

    except Exception as e:
        return jsonify({'success': False, 'message': f'Lỗi hệ thống: {str(e)}'}), 500