"""
초고화질 이미지 자동 추출기 v6.1 - 직접 추출 방식
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
- xlsx 내부 이미지를 직접 추출 (원본 해상도 100%)
- 열(column) 위치 기반 상면/측면 정확 분류
  · 홀수 열(1,3,5,7) + 큰 이미지(disp_h≥60) → 상면
  · 짝수 열(2,4,6,8) 또는 홀수 열 납작 이미지(disp_h<60) → 측면
- .xls 파일: 같은 폴더 xlsx 우선 사용, 없으면 COM 변환 시도
"""

import os, sys, io, re, time
from pathlib import Path
from PIL import Image

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)
import cutting_eval_tool as engine


# ─── 유틸 ───────────────────────────────────────────────────
def safe_name(text):
    return "".join(c for c in str(text) if c.isalnum() or c in " _-+").strip() or "UNNAMED"


def save_image(raw_bytes, save_path):
    """srcRect 크롭 없이 원본 이미지 그대로 JPEG 고화질 저장 (품질 97%)"""
    im = Image.open(io.BytesIO(raw_bytes)).convert('RGB')
    im.save(str(save_path), "JPEG", quality=97, subsampling=0)
    return im.size


# ─── .xls → .xlsx 경로 확보 ──────────────────────────────────
def resolve_xlsx(file_path):
    """
    입력 파일에서 사용할 xlsx 경로를 결정.
    1. 입력이 xlsx → 그대로 반환
    2. 입력이 xls  → 같은 폴더 / _data 하위 폴더 / 상위 _data 폴더 순으로 동명 xlsx 탐색
    3. 없으면 COM(DispatchEx)으로 변환 시도
    4. 변환도 실패하면 None 반환
    """
    abs_path = os.path.abspath(file_path)
    ext = os.path.splitext(abs_path)[1].lower()

    if ext == '.xlsx':
        return abs_path

    base_name = os.path.splitext(os.path.basename(abs_path))[0] + '.xlsx'
    parent    = os.path.dirname(abs_path)

    # 탐색 순서: 같은 폴더 → _data 하위 → 상위 폴더 → 상위/_data
    search_dirs = [
        parent,
        os.path.join(parent, '_data'),
        os.path.dirname(parent),
        os.path.join(os.path.dirname(parent), '_data'),
    ]
    for d in search_dirs:
        candidate = os.path.join(d, base_name)
        if os.path.exists(candidate):
            print(f"  xlsx 발견: {candidate}")
            return candidate

    # COM으로 변환 시도
    print("  .xls → xlsx 변환 중 (Excel COM)...")
    xlsx_out = os.path.join(parent, base_name)
    excel = None
    try:
        import win32com.client, pythoncom
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(abs_path)
        wb.SaveAs(xlsx_out, FileFormat=51)
        wb.Close(SaveChanges=False)
        print(f"  변환 완료: {os.path.basename(xlsx_out)}")
        return xlsx_out
    except Exception as e:
        print(f"  [변환 실패] {e}")
        return None
    finally:
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass


# ─── 핵심 로직 ──────────────────────────────────────────────
def run_extraction(file_path):
    print("=" * 65)
    print("  [초고화질 이미지 자동 추출기 v6.1]")
    print(f"  대상 파일: {os.path.basename(file_path)}")
    print("=" * 65)

    abs_path = os.path.abspath(file_path)
    ext = os.path.splitext(abs_path)[1].lower()

    try:
        # ── 1. xlsx 경로 확보 ────────────────────────────────
        print("> 파일 확인 중...")
        xlsx_path = resolve_xlsx(abs_path)
        if xlsx_path is None:
            print("\n[중단] xlsx를 구할 수 없습니다.")
            print("  해결: 해당 xls를 Excel에서 xlsx로 저장 후 같은 폴더에 두세요.")
            input("\n아무 키나 누르면 종료됩니다.")
            return

        # ── 2. 데이터 파싱 ───────────────────────────────────
        print("> 셀 데이터 파싱 중...")
        if ext == '.xls':
            row_h, col_w, text_cells, merges, n_rows, n_cols = engine.parse_xls(abs_path)
        else:
            row_h, col_w, text_cells, merges, n_rows, n_cols = engine.parse_xls_from_xlsx(xlsx_path)

        print("> 이미지 데이터 추출 중...")
        img_cells, img_data = engine.parse_images(xlsx_path)
        print(f"  [발견] 내장 이미지 {len(img_data)}개, 이미지 셀 {len(img_cells)}개")

        if not img_data:
            print("  [경고] 추출할 이미지가 없습니다.")
            input("\n아무 키나 누르면 종료됩니다.")
            return

        # ── 3. 출력 폴더 생성 ────────────────────────────────
        output_dir = Path(abs_path).parent / f"추출결과_{Path(abs_path).stem}"
        output_dir.mkdir(exist_ok=True)

        # ── 4. 패스(Pass) 행 파악 ────────────────────────────
        #   A열(col=0)에서 "숫자P" 패턴 탐색
        pass_rows = {}
        for (r, c), txt in text_cells.items():
            if c == 0:
                m = re.match(r'(\d+)\s*P', txt.strip(), re.IGNORECASE)
                if m:
                    pass_rows[r] = int(m.group(1))

        sorted_pr = sorted(pass_rows.keys())
        pass_ranges = {}
        for i, pr in enumerate(sorted_pr):
            pn = pass_rows[pr]
            end = sorted_pr[i + 1] if i + 1 < len(sorted_pr) else n_rows
            pass_ranges[pn] = (pr, end)

        if not pass_ranges:
            print("  [정보] 패스 정보 없음 → 전체를 단일 블록으로 처리")
            pass_ranges[1] = (0, n_rows)

        # ── 5. 인서트(Grade) 열 파악 ─────────────────────────
        #   홀수 열(1,3,5,7,...)이 각 인서트 시작 열
        #   row=1: 번호, row=2: 인서트형, row=3: Grade
        insert_cols = {}
        for (r, c), txt in text_cells.items():
            if c % 2 == 1 and c > 0:
                if c not in insert_cols:
                    insert_cols[c] = {}
                if r == 1:
                    insert_cols[c]['num'] = txt
                elif r == 2:
                    insert_cols[c]['chip'] = txt
                elif r == 3:
                    insert_cols[c]['grade'] = txt

        if not insert_cols:
            # 이미지가 있는 홀수 열을 추론
            odd_img_cols = sorted(set(c for (r, c) in img_cells if c % 2 == 1))
            for fc in odd_img_cols:
                insert_cols[fc] = {'grade': f'COL{fc}'}

        print(f"> 추출 시작 (패스 {len(pass_ranges)}단계, 인서트 {len(insert_cols)}종)")
        count = 0

        # ── 6. 패스 × 인서트 순회 ────────────────────────────
        for pn, (rs, re_row) in sorted(pass_ranges.items()):
            for col_start in sorted(insert_cols.keys()):
                grade   = insert_cols[col_start].get('grade', f'COL{col_start}')
                chip    = insert_cols[col_start].get('chip', '')
                num     = insert_cols[col_start].get('num', str(col_start))

                # 인서트 1종 = 홀수 열(col_start) + 짝수 열(col_start+1)
                odd_col  = col_start          # 상면 주 열
                even_col = col_start + 1      # 측면 주 열

                # 패스 구간 내 이미지 수집
                top_candidates      = []  # 상면 후보: 홀수 열 큰 이미지 (disp_h ≥ 60)
                side_odd            = []  # 측면 1순위: 홀수 열 납작 이미지 (0P와 동일 형태)
                side_even           = []  # 측면 2순위: 짝수 열 이미지 (fallback)

                for r in range(rs, re_row):
                    if (r, odd_col) in img_cells:
                        for inf in img_cells[(r, odd_col)]:
                            if inf.get('is_r_part'):    # 썸네일/오버레이 제외
                                continue
                            if inf.get('disp_h', 0) >= 60:
                                top_candidates.append(inf)
                            else:
                                side_odd.append(inf)    # 납작한 이미지 → 측면 (0P와 동일)
                    if (r, even_col) in img_cells:
                        for inf in img_cells[(r, even_col)]:
                            side_even.append(inf)

                # 측면: 홀수 열 납작 이미지 우선 (0P와 같은 형태), 없으면 짝수 열 사용
                side_candidates = side_odd if side_odd else side_even

                if not top_candidates and not side_candidates:
                    continue

                # ── 7. 대표 이미지 선택 ──────────────────────
                #   is_crater = 겹침 셀의 큰 배경 이미지 (실제 인서트 사진)
                #   is_r_part = 겹침 셀의 작은 오버레이 → 제외
                def pick_best(pool):
                    if not pool:
                        return None
                    # is_crater 이미지가 있으면 최우선 (실제 인서트 상면 사진)
                    craters = [x for x in pool if x.get('is_crater') and not x.get('is_r_part')]
                    if craters:
                        return max(craters, key=lambda x: x.get('disp_w', 0) * x.get('disp_h', 0))
                    # 없으면 r_part 제외 후 가장 큰 것
                    filtered = [x for x in pool if not x.get('is_r_part')]
                    return max(filtered or pool,
                               key=lambda x: x.get('disp_w', 0) * x.get('disp_h', 0))

                # ── 8. 폴더 구조: 추출결과/[Grade명]/상면|측면/NP.png ─
                folder_name = safe_name(grade) or safe_name(chip) or f"인서트{num}"
                tool_dir = output_dir / folder_name
                (tool_dir / "상면").mkdir(parents=True, exist_ok=True)
                (tool_dir / "측면").mkdir(parents=True, exist_ok=True)

                pass_str = f"{pn}P"

                for pool, subfolder in [(top_candidates, "상면"),
                                        (side_candidates, "측면")]:
                    best = pick_best(pool)
                    if best is None:
                        continue
                    raw = img_data.get(best['fname'])
                    if not raw:
                        print(f"    ! [{subfolder}] {pass_str} 데이터 없음 ({best['fname']})")
                        continue
                    save_path = tool_dir / subfolder / f"{pass_str}.jpg"
                    try:
                        w, h = save_image(raw, save_path)
                        count += 1
                        print(f"    + [{subfolder}] {folder_name} / {pass_str}.png  ({w}×{h}px)")
                    except Exception as e:
                        print(f"    ! [{subfolder}] {pass_str} 저장 실패: {e}")

        # ── 9. 완료 ──────────────────────────────────────────
        print("=" * 65)
        print(f"  [완료] 총 {count}장 추출 (원본 화질 100%)")
        print(f"  저장 경로: {output_dir}")
        print("=" * 65)
        os.startfile(str(output_dir))

    except Exception as e:
        print(f"\n[오류 발생] {e}")
        import traceback
        traceback.print_exc()

    print("\n아무 키나 누르면 종료됩니다.")
    input()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("사용법: python cli_extractor.py <엑셀파일경로>")
        time.sleep(3)
    else:
        run_extraction(sys.argv[1])
