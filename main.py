"""
나라장터 입찰공고 키워드 알림
- 공공데이터포털 나라장터 입찰공고정보서비스에서 최근 공고를 가져와
- 설정한 키워드가 공고명에 포함된 신규 공고만 골라
- 구글 시트에 기록하고 Gmail로 알림을 보냅니다.

민감 정보는 모두 환경변수(GitHub Secrets)에서 읽습니다:
  DATA_API_KEY       : 공공데이터포털 인증키 (Decoding 키)
  GOOGLE_CREDENTIALS : 서비스 계정 JSON 파일 내용 전체
  GMAIL_FROM         : 보내는 Gmail 주소
  GMAIL_PASSWORD     : Gmail 앱 비밀번호 (16자리, 띄어쓰기 없이)
  GMAIL_TO           : 받는 이메일 주소 (쉼표로 여러 개 가능)
"""

import os
import json
import smtplib
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import requests
import gspread
from google.oauth2.service_account import Credentials


# ---------------------------------------------------------------------------
# 설정 로드
# ---------------------------------------------------------------------------
def load_config():
    with open("config.json", "r", encoding="utf-8") as f:
        return json.load(f)


def get_env(name):
    value = os.environ.get(name)
    if not value:
        raise RuntimeError(f"환경변수 {name} 가(이) 설정되지 않았습니다.")
    return value


# ---------------------------------------------------------------------------
# 나라장터 API 호출
# ---------------------------------------------------------------------------
API_BASE = "https://apis.data.go.kr/1230000/ad/BidPublicInfoService"

# 업무 구분별 오퍼레이션
OPERATIONS = {
    "용역": "getBidPblancListInfoServc",
    "물품": "getBidPblancListInfoThng",
    "공사": "getBidPblancListInfoCnstwk",
    "외자": "getBidPblancListInfoFrgcpt",
}


def fetch_bids(api_key, operation, begin_dt, end_dt, rows):
    """한 업무 구분의 입찰공고 목록을 가져옵니다."""
    url = f"{API_BASE}/{operation}"
    params = {
        "serviceKey": api_key,
        "pageNo": "1",
        "numOfRows": str(rows),
        "inqryDiv": "1",          # 1 = 공고게시일시 기준
        "inqryBgnDt": begin_dt,   # YYYYMMDDHHMM
        "inqryEndDt": end_dt,
        "type": "json",
    }
    resp = requests.get(url, params=params, timeout=30)
    resp.raise_for_status()
    data = resp.json()

    # 정상 응답 확인
    header = data.get("response", {}).get("header", {})
    if header.get("resultCode") not in ("00", "0"):
        print(f"  [{operation}] API 응답 이상: {header.get('resultMsg')}")
        return []

    body = data.get("response", {}).get("body", {})
    items = body.get("items", [])
    # items 가 단일 dict 로 올 때도 있어 리스트로 정규화
    if isinstance(items, dict):
        items = [items]
    return items


# ---------------------------------------------------------------------------
# 구글 시트
# ---------------------------------------------------------------------------
def open_worksheet(config):
    creds_json = json.loads(get_env("GOOGLE_CREDENTIALS"))
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_json, scopes=scopes)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(config["sheet_id"])
    return spreadsheet.worksheet(config["worksheet_name"])


def get_existing_ids(worksheet):
    """이미 시트에 기록된 공고번호 집합 (A열)."""
    col = worksheet.col_values(1)  # A열 전체
    # 첫 행은 제목이므로 제외
    return set(col[1:]) if len(col) > 1 else set()


# ---------------------------------------------------------------------------
# 이메일
# ---------------------------------------------------------------------------
def send_email(new_items):
    gmail_from = get_env("GMAIL_FROM")
    gmail_pw = get_env("GMAIL_PASSWORD")
    gmail_to = [addr.strip() for addr in get_env("GMAIL_TO").split(",")]

    subject = f"[나라장터 알림] 신규 공고 {len(new_items)}건"

    # 본문 (HTML)
    rows_html = ""
    for it in new_items:
        rows_html += f"""
        <tr>
          <td style="padding:8px;border:1px solid #ddd;">{it['업무']}</td>
          <td style="padding:8px;border:1px solid #ddd;">{it['공고명']}</td>
          <td style="padding:8px;border:1px solid #ddd;">{it['기관']}</td>
          <td style="padding:8px;border:1px solid #ddd;white-space:nowrap;">{it['입찰마감일'] or '-'}</td>
          <td style="padding:8px;border:1px solid #ddd;">
            {'<a href="' + it['상세링크'] + '">바로가기</a>' if it['상세링크'] else '-'}
          </td>
        </tr>"""

    html = f"""
    <div style="font-family:sans-serif;">
      <h2>나라장터 신규 입찰공고 {len(new_items)}건</h2>
      <p>설정하신 키워드에 해당하는 신규 공고가 등록되었습니다.</p>
      <table style="border-collapse:collapse;width:100%;font-size:14px;">
        <thead>
          <tr style="background:#f2f2f2;">
            <th style="padding:8px;border:1px solid #ddd;">업무</th>
            <th style="padding:8px;border:1px solid #ddd;">공고명</th>
            <th style="padding:8px;border:1px solid #ddd;">기관</th>
            <th style="padding:8px;border:1px solid #ddd;">입찰마감</th>
            <th style="padding:8px;border:1px solid #ddd;">링크</th>
          </tr>
        </thead>
        <tbody>{rows_html}</tbody>
      </table>
      <p style="color:#888;font-size:12px;margin-top:16px;">
        자세한 내용은 연결된 구글 시트에서도 확인할 수 있습니다.
      </p>
    </div>"""

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = gmail_from
    msg["To"] = ", ".join(gmail_to)
    msg.attach(MIMEText(html, "html"))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(gmail_from, gmail_pw)
        server.sendmail(gmail_from, gmail_to, msg.as_string())

    print(f"이메일 발송 완료 → {', '.join(gmail_to)}")


# ---------------------------------------------------------------------------
# 메인
# ---------------------------------------------------------------------------
def main():
    config = load_config()
    api_key = get_env("DATA_API_KEY")
    keywords = config["keywords"]

    # 조회 기간 계산 (최근 N일)
    now = datetime.now()
    begin = now - timedelta(days=config["inquiry_days"])
    begin_dt = begin.strftime("%Y%m%d%H%M")
    end_dt = now.strftime("%Y%m%d%H%M")
    print(f"조회 기간: {begin_dt} ~ {end_dt}")
    print(f"키워드: {', '.join(keywords)}")

    # 구글 시트 열기 + 기존 공고번호 읽기
    worksheet = open_worksheet(config)
    existing_ids = get_existing_ids(worksheet)
    print(f"시트에 기록된 기존 공고: {len(existing_ids)}건")

    # 업무별로 API 호출 → 키워드 필터 → 신규만 수집
    new_items = []
    new_rows = []
    for work_type, operation in OPERATIONS.items():
        items = fetch_bids(api_key, operation, begin_dt, end_dt,
                           config["rows_per_request"])
        print(f"  [{work_type}] {len(items)}건 조회")

        for it in items:
            title = it.get("bidNtceNm", "")
            bid_no = it.get("bidNtceNo", "")

            # 키워드 포함 여부
            if not any(kw in title for kw in keywords):
                continue
            # 중복 체크
            if bid_no in existing_ids:
                continue

            record = {
                "공고번호": bid_no,
                "공고명": title,
                "기관": it.get("ntceInsttNm", ""),
                "업무": work_type,
                "공고일시": it.get("bidNtceDt", ""),
                "입찰마감일": it.get("bidClseDt", ""),
                "개찰일시": it.get("opengDt", ""),
                "상세링크": it.get("ntceSpecDocUrl1", ""),
            }
            new_items.append(record)
            new_rows.append([
                record["공고번호"], record["공고명"], record["기관"],
                record["업무"], record["공고일시"], record["입찰마감일"],
                record["개찰일시"], record["상세링크"],
            ])
            existing_ids.add(bid_no)  # 같은 실행 내 중복도 방지

    print(f"신규 매칭 공고: {len(new_items)}건")

    if not new_items:
        print("신규 공고가 없습니다. 종료합니다.")
        return

    # 시트에 추가
    worksheet.append_rows(new_rows, value_input_option="USER_ENTERED")
    print(f"시트에 {len(new_rows)}건 기록 완료")

    # 이메일 발송
    send_email(new_items)


if __name__ == "__main__":
    main()
