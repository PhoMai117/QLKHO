import os
import threading
import webbrowser
import http.server
import socketserver
from flask import Flask
from flask_cors import CORS

# Import Module Đăng nhập (Blueprint)
from modules.auth import auth_bp
from modules.production import production_bp
from modules.warehouse import warehouse_bp

app = Flask(__name__)
# Cho phép Frontend ở Port 4000 gọi API sang Backend ở Port 5000
CORS(app)

# Đăng ký Blueprint
app.register_blueprint(auth_bp, url_prefix='/api/auth')
app.register_blueprint(production_bp, url_prefix='/api/production')
app.register_blueprint(warehouse_bp, url_prefix='/api/warehouse')
# ==========================================
# CẤU HÌNH SERVER FRONTEND NGẦM (PORT 4000)
# ==========================================
def chay_frontend_ngam():
    # Trỏ về thư mục /Front
    front_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'Front'))
    
    if os.path.exists(front_dir):
        os.chdir(front_dir)
    
    PORT = 4000
    Handler = http.server.SimpleHTTPRequestHandler
    socketserver.TCPServer.allow_reuse_address = True
    
    with socketserver.TCPServer(("", PORT), Handler) as httpd:
        print(f"-> Frontend Server đang chạy tại: http://localhost:{PORT}")
        httpd.serve_forever()

if __name__ == '__main__':
    # 1. Chạy luồng Frontend ở cổng 4000
    frontend_thread = threading.Thread(target=chay_frontend_ngam, daemon=True)
    frontend_thread.start()
    
    # 2. Tự động bật trình duyệt mở trang login sau 1.5s
    threading.Timer(1.5, lambda: webbrowser.open('http://localhost:4000/login.html')).start()

    # 3. Khởi chạy Backend Flask ở cổng 5000
    print("-> Backend Flask đang khởi động tại: http://localhost:5000")
    app.run(host='0.0.0.0', port=5000, debug=False)