"""
컴퓨터활용능력 2급 실기 자동 채점 프로그램 (핵심 채점 모듈)
- 정답.xlsx 파일과 제출.xlsx 파일을 비교하여
  1) 지정 범위(예: B2:D10)의 '값' 일치 여부
  2) 수식 범위(예: E2:E10)의 '결과값' + '사용한 함수 이름' 일치 여부
  를 검사하고, 틀린 셀 위치와 100점 만점 기준 점수를 계산합니다.
"""

import re
import math
from pathlib import Path
import openpyxl


def normalize_value(val):
    """
    셀 값을 비교하기 전에 형식을 통일하는 함수입니다.
    - 빈 셀(None)과 빈 문자열("")을 똑같이 취급합니다.
    - 문자열 앞뒤의 불필요한 공백을 제거합니다.
    """
    if val is None:
        return ""
    if isinstance(val, str):
        return val.strip()
    return val


def values_match(ans_val, sub_val):
    """
    두 셀의 '결과값'이 같은지 비교합니다.
    숫자(정수/실수)는 소수점 오차를 감안해 비교합니다. (예: 10 == 10.0)
    """
    a = normalize_value(ans_val)
    b = normalize_value(sub_val)

    # 둘 다 숫자(int 또는 float, 단 bool 제외)인 경우
    if isinstance(a, (int, float)) and not isinstance(a, bool) and \
       isinstance(b, (int, float)) and not isinstance(b, bool):
        return math.isclose(float(a), float(b), rel_tol=1e-9, abs_tol=1e-9)

    return a == b


def extract_function_names(formula_text):
    """
    엑셀 수식 문자열에서 사용한 함수 이름 목록만 뽑아냅니다.
    예: '=IF(AVERAGE(B2:D2)>=70, "합격", "불합격")'
        -> ['AVERAGE', 'IF'] (정렬된 목록으로 반환)
    """
    if not isinstance(formula_text, str) or not formula_text.startswith("="):
        return []

    # 영문/숫자/점(.)으로 이루어진 이름 바로 뒤에 '('가 오는 패턴을 찾습니다.
    names = re.findall(r"([A-Z0-9_.]+)\s*\(", formula_text.upper())
    # 순서에 상관없이 어떤 함수들을 썼는지 비교하기 위해 집합(중복 제거) 후 정렬합니다.
    return sorted(set(names))


def grade_excel(
    answer_path: str,
    submit_path: str,
    sheet_name: str = None,
    value_range: str = "B2:D10",
    formula_range: str = "E2:E10",
):
    """
    정답 엑셀 파일과 제출 엑셀 파일을 비교 채점하는 메인 함수입니다.

    매개변수:
    - answer_path: 정답 엑셀 파일 경로 (예: '정답.xlsx')
    - submit_path: 학생이 제출한 엑셀 파일 경로 (예: '제출.xlsx')
    - sheet_name: 비교할 시트 이름 (None이면 첫 번째 시트를 자동으로 선택)
    - value_range: 값만 비교할 셀 범위 (예: 'B2:D10')
    - formula_range: 결과값과 함수 이름을 함께 비교할 셀 범위 (예: 'E2:E10')
    """

    # 1. 엑셀 파일 열기
    # openpyxl은 파일을 두 가지 모드로 열 수 있습니다:
    # - data_only=True  : 수식이 계산된 '결과값'을 읽어옵니다.
    # - data_only=False : 셀 안에 적힌 '=SUM(B2:D2)' 같은 '수식 글자'를 읽어옵니다.
    wb_ans_val = openpyxl.load_workbook(answer_path, data_only=True)
    wb_ans_f = openpyxl.load_workbook(answer_path, data_only=False)

    wb_sub_val = openpyxl.load_workbook(submit_path, data_only=True)
    wb_sub_f = openpyxl.load_workbook(submit_path, data_only=False)

    # 2. 채점할 시트 선택하기
    target_sheet = sheet_name or wb_ans_val.sheetnames[0]
    if target_sheet not in wb_sub_val.sheetnames:
        return {
            "score": 0.0,
            "total_cells": 0,
            "correct_cells": 0,
            "errors": [
                {
                    "cell": "시트 전체",
                    "type": "시트 누락",
                    "detail": f"제출 파일에 '{target_sheet}' 시트가 없습니다.",
                }
            ],
        }

    ws_ans_val = wb_ans_val[target_sheet]
    ws_ans_f = wb_ans_f[target_sheet]
    ws_sub_val = wb_sub_val[target_sheet]
    ws_sub_f = wb_sub_f[target_sheet]

    total_cells = 0    # 전체 채점 대상 셀 개수
    correct_cells = 0  # 맞힌 셀 개수
    errors = []        # 틀린 셀 정보를 담을 리스트

    # 3. [값 비교] 지정된 범위(예: B2:D10)의 값을 한 칸씩 비교
    for row_ans, row_sub in zip(ws_ans_val[value_range], ws_sub_val[value_range]):
        for cell_ans, cell_sub in zip(row_ans, row_sub):
            total_cells += 1
            coord = cell_ans.coordinate  # 셀 위치 (예: 'C5')

            if values_match(cell_ans.value, cell_sub.value):
                correct_cells += 1
            else:
                errors.append(
                    {
                        "cell": coord,
                        "type": "값 불일치",
                        "detail": (
                            f"{coord} 셀: 정답은 '{normalize_value(cell_ans.value)}'인데 "
                            f"제출값은 '{normalize_value(cell_sub.value)}'입니다."
                        ),
                    }
                )

    # 4. [수식 + 결과값 비교] 지정된 범위(예: E2:E10) 검사
    for row_ans_v, row_sub_v, row_ans_f, row_sub_f in zip(
        ws_ans_val[formula_range],
        ws_sub_val[formula_range],
        ws_ans_f[formula_range],
        ws_sub_f[formula_range],
    ):
        for c_ans_v, c_sub_v, c_ans_f, c_sub_f in zip(
            row_ans_v, row_sub_v, row_ans_f, row_sub_f
        ):
            total_cells += 1
            coord = c_ans_v.coordinate

            ans_funcs = extract_function_names(c_ans_f.value)
            sub_funcs = extract_function_names(c_sub_f.value)

            # (1) 정답에 함수가 쓰였는데 제출 파일에 수식/함수가 없는 경우
            if ans_funcs and not sub_funcs:
                errors.append(
                    {
                        "cell": coord,
                        "type": "수식 미사용",
                        "detail": (
                            f"{coord} 셀: 함수({', '.join(ans_funcs)})를 사용하지 않고 "
                            f"직접 값('{normalize_value(c_sub_v.value)}')만 입력했습니다."
                        ),
                    }
                )
                continue

            # (2) 사용한 함수 이름이 정답과 다른 경우 (예: SUM 대신 AVERAGE 사용)
            if ans_funcs != sub_funcs:
                errors.append(
                    {
                        "cell": coord,
                        "type": "함수 이름 불일치",
                        "detail": (
                            f"{coord} 셀: 정답 함수는 {ans_funcs}이지만, "
                            f"제출 파일은 {sub_funcs} 함수를 사용했습니다."
                        ),
                    }
                )
                continue

            # (3) 함수 이름은 맞지만 계산 결과값이 다른 경우 (예: 셀 범위 지정 실수)
            # 참고: openpyxl로 새로 만든 파일처럼 캐시된 결과값이 None일 때는 수식 문자열도 보조 비교
            if c_ans_v.value is not None and c_sub_v.value is not None:
                if not values_match(c_ans_v.value, c_sub_v.value):
                    errors.append(
                        {
                            "cell": coord,
                            "type": "수식 결과값 불일치",
                            "detail": (
                                f"{coord} 셀: 함수({', '.join(sub_funcs)})는 맞았으나 "
                                f"결과값이 다릅니다. (정답: '{c_ans_v.value}', 제출: '{c_sub_v.value}')"
                            ),
                        }
                    )
                    continue
            elif str(c_ans_f.value).upper().replace(" ", "") != str(c_sub_f.value).upper().replace(" ", ""):
                errors.append(
                    {
                        "cell": coord,
                        "type": "수식 범위/인수 불일치",
                        "detail": (
                            f"{coord} 셀: 수식 구성이 정답과 다릅니다. "
                            f"(정답: '{c_ans_f.value}', 제출: '{c_sub_f.value}')"
                        ),
                    }
                )
                continue

            # 모든 검사를 통과하면 정답 처리!
            correct_cells += 1

    # 5. 총 100점 만점 기준 점수 계산
    score = round((correct_cells / total_cells) * 100, 1) if total_cells > 0 else 0.0

    return {
        "score": score,
        "total_cells": total_cells,
        "correct_cells": correct_cells,
        "errors": errors,
    }


def print_report(result):
    """채점 결과를 보기 좋게 화면에 출력하는 함수입니다."""
    print("=" * 55)
    print("        컴퓨터활용능력 2급 실기 채점 결과표")
    print("=" * 55)
    print(f"총점: {result['score']}점 / 100점")
    print(f"정답 셀: {result['correct_cells']}개 / 전체 {result['total_cells']}개")
    print("-" * 55)

    if not result["errors"]:
        print("🎉 감점 내역이 없습니다! 완벽한 만점입니다.")
    else:
        print(f"틀린 내역 ({len(result['errors'])}건):")
        for idx, err in enumerate(result["errors"], start=1):
            print(f"  {idx}. [{err['type']}] {err['detail']}")
    print("=" * 55)


if __name__ == "__main__":
    # 직접 실행할 때 현재 폴더의 정답.xlsx / 제출.xlsx를 찾아 채점합니다.
    ans_file = Path("정답.xlsx")
    sub_file = Path("제출.xlsx")

    if ans_file.exists() and sub_file.exists():
        res = grade_excel(str(ans_file), str(sub_file))
        print_report(res)
    else:
        print("안내: 같은 폴더에 '정답.xlsx'와 '제출.xlsx' 파일을 두고 실행해 주세요.")
