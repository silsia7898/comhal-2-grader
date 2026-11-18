"""
컴퓨터활용능력 자동 채점기 웹 서버
- Python 기본 내장 모듈(http.server, json, base64, tempfile)만 사용합니다.
- 정적 HTML 화면(auto_grader.html)을 서빙하고,
  POST /api/grade 요청을 받아 grade_excel.py의 grade_excel() 함수를 실행합니다.
"""

import base64
import json
import os
import tempfile
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from grade_excel import grade_excel

BASE_DIR = Path(__file__).resolve().parent


class GraderRequestHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(BASE_DIR), **kwargs)

    def do_GET(self):
        # 기본 경로('/') 접속 시 auto_grader.html을 보여줍니다.
        if self.path == "/" or self.path == "":
            self.path = "/auto_grader.html"
        return super().do_GET()

    def do_POST(self):
        if self.path != "/api/grade":
            self.send_json(404, {"error": "지원하지 않는 API 경로입니다."})
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0"))
            raw_body = self.rfile.read(content_length)
            payload = json.loads(raw_body.decode("utf-8"))

            answer_b64 = payload.get("answerBase64", "")
            submit_b64 = payload.get("submitBase64", "")
            value_range = (payload.get("valueRange") or "B2:D10").strip()
            formula_range = (payload.get("formulaRange") or "E2:E10").strip()
            sheet_name = (payload.get("sheetName") or "").strip() or None

            if not answer_b64 or not submit_b64:
                self.send_json(400, {"error": "파일 두 개를 모두 올려주세요!"})
                return

            answer_bytes = base64.b64decode(answer_b64)
            submit_bytes = base64.b64decode(submit_b64)

            # 임시 폴더에 두 엑셀 파일을 저장한 뒤 grade_excel()로 비교 채점합니다.
            with tempfile.TemporaryDirectory() as tmp_dir:
                ans_path = os.path.join(tmp_dir, "정답.xlsx")
                sub_path = os.path.join(tmp_dir, "제출.xlsx")

                with open(ans_path, "wb") as f:
                    f.write(answer_bytes)
                with open(sub_path, "wb") as f:
                    f.write(submit_bytes)

                result = grade_excel(
                    answer_path=ans_path,
                    submit_path=sub_path,
                    sheet_name=sheet_name,
                    value_range=value_range,
                    formula_range=formula_range,
                )

            self.send_json(200, result)

        except Exception as exc:
            self.send_json(500, {"error": f"채점 중 오류가 발생했습니다: {exc}"})

    def send_json(self, status_code: int, data: dict):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def run_server(port: int = 8000):
    server = ThreadingHTTPServer(("127.0.0.1", port), GraderRequestHandler)
    print(f"컴퓨터활용능력 자동 채점기 서버 실행 중: http://127.0.0.1:{port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n서버를 종료합니다.")
        server.server_close()


if __name__ == "__main__":
    run_server(8000)
