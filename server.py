# run_server.py

from waitress import serve
from app import app  # 'app'은 Flask 애플리케이션이 정의된 모듈과 인스턴스 이름입니다

if __name__ == "__main__":
    serve(app, host="0.0.0.0", port=5000)
