import os
import json
import re
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CLIENT_SECRETS = os.path.join(BASE_DIR, 'client_secrets.json.json')
TOKEN_FILE = os.path.join(BASE_DIR, 'token_sheets.json')  # Sheets 전용 토큰
DATA_JS = os.path.join(BASE_DIR, 'data', 'data.js')

import os
from dotenv import load_dotenv
load_dotenv()
SPREADSHEET_ID = os.environ.get('SPREADSHEET_ID')
SHEET_NAME = '26년 홍보 컨텐츠 발행내역'
DATA_START_ROW = 9  # 8행이 헤더, 9행부터 데이터


def get_sheets_service():
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request

    # GitHub Actions 환경: 환경변수로 인증
    refresh_token = os.environ.get('SHEETS_REFRESH_TOKEN')
    client_id = os.environ.get('SHEETS_CLIENT_ID')
    client_secret = os.environ.get('SHEETS_CLIENT_SECRET')

    if refresh_token and client_id and client_secret:
        credentials = Credentials(
            token=None,
            refresh_token=refresh_token,
            token_uri='https://oauth2.googleapis.com/token',
            client_id=client_id,
            client_secret=client_secret,
            scopes=SCOPES
        )
        credentials.refresh(Request())
        return build('sheets', 'v4', credentials=credentials)

    # 로컬 환경: token_sheets.json 사용
    credentials = None
    if os.path.exists(TOKEN_FILE):
        credentials = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            try:
                credentials.refresh(Request())
            except Exception:
                credentials = None

        if not credentials or not credentials.valid:
            print("Google 인증이 필요합니다. 브라우저가 열립니다...")
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRETS, SCOPES)
            credentials = flow.run_local_server(port=0)
            with open(TOKEN_FILE, 'w') as f:
                f.write(credentials.to_json())

    return build('sheets', 'v4', credentials=credentials)


def load_data_js():
    with open(DATA_JS, 'r', encoding='utf-8') as f:
        content = f.read()
    json_str = re.sub(r'^const metricsData\s*=\s*', '', content.strip()).rstrip(';').strip()
    return json.loads(json_str)


def get_existing_data(service):
    """시트에서 기존 데이터 전체 읽기 (행 번호 포함)"""
    range_name = f"'{SHEET_NAME}'!B{DATA_START_ROW}:F"
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=range_name
    ).execute()
    rows = result.get('values', [])
    existing = []
    for i, row in enumerate(rows):
        title = row[1].strip() if len(row) > 1 else ''
        date = row[2].strip() if len(row) > 2 else ''
        channel = row[3].strip() if len(row) > 3 else ''
        if title:
            existing.append({
                'title': title,
                'date': date,
                'channel': channel,
                'row_index': DATA_START_ROW + i  # 실제 행 번호
            })
    return existing


def delete_rows(service, sheet_id, row_indices):
    """지정한 행들 삭제 (역순으로 삭제해야 인덱스 유지)"""
    requests = []
    for row_index in sorted(row_indices, reverse=True):
        requests.append({
            'deleteDimension': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'ROWS',
                    'startIndex': row_index - 1,  # 0-based
                    'endIndex': row_index
                }
            }
        })
    if requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={'requests': requests}
        ).execute()


def get_last_row(service):
    """데이터가 있는 마지막 행 번호"""
    range_name = f"'{SHEET_NAME}'!B{DATA_START_ROW}:B"
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=range_name
    ).execute()
    rows = result.get('values', [])
    return DATA_START_ROW + len(rows) - 1


def get_sheet_id(service):
    """시트 GID 조회"""
    spreadsheet = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    for sheet in spreadsheet['sheets']:
        if sheet['properties']['title'] == SHEET_NAME:
            return sheet['properties']['sheetId']
    return 0


def append_rows(service, rows):
    """새 행 추가 + B열 수식 + 서식 복사"""
    last_row = get_last_row(service)
    new_start = last_row + 1
    sheet_id = get_sheet_id(service)

    formatted_rows = []
    for i, row in enumerate(rows):
        row_num = new_start + i
        new_row = [f'=MONTH(D{row_num})'] + list(row)[1:]
        formatted_rows.append(new_row)

    service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{SHEET_NAME}'!B:G",
        valueInputOption='USER_ENTERED',
        insertDataOption='INSERT_ROWS',
        body={'values': formatted_rows}
    ).execute()

    new_count = len(rows)
    total_last_row = new_start - 1 + new_count

    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={'requests': [
            {
                'copyPaste': {
                    'source': {
                        'sheetId': sheet_id,
                        'startRowIndex': DATA_START_ROW - 1,
                        'endRowIndex': DATA_START_ROW,
                        'startColumnIndex': 1,
                        'endColumnIndex': 7
                    },
                    'destination': {
                        'sheetId': sheet_id,
                        'startRowIndex': new_start - 1,
                        'endRowIndex': total_last_row,
                        'startColumnIndex': 1,
                        'endColumnIndex': 7
                    },
                    'pasteType': 'PASTE_FORMAT'
                }
            },
            {
                'setBasicFilter': {
                    'filter': {
                        'range': {
                            'sheetId': sheet_id,
                            'startRowIndex': DATA_START_ROW - 2,
                            'endRowIndex': total_last_row,
                            'startColumnIndex': 1,
                            'endColumnIndex': 7
                        }
                    }
                }
            }
        ]}
    ).execute()


def update_subscriber_counts(service, data):
    """E4(링크드인 팔로워), G4(유튜브 구독자) 업데이트"""
    history = sorted(data.get('history', []), key=lambda x: x['date'])

    linkedin_followers = None
    youtube_subscribers = None
    for h in reversed(history):
        if linkedin_followers is None and h.get('linkedin_followers', 0) > 0:
            linkedin_followers = h['linkedin_followers']
        if youtube_subscribers is None and h.get('youtube_subscribers', 0) > 0:
            youtube_subscribers = h['youtube_subscribers']
        if linkedin_followers is not None and youtube_subscribers is not None:
            break

    updates = []
    if linkedin_followers is not None:
        updates.append({'range': f"'{SHEET_NAME}'!F4", 'values': [[linkedin_followers]]})
    if youtube_subscribers is not None:
        updates.append({'range': f"'{SHEET_NAME}'!G4", 'values': [[youtube_subscribers]]})

    if updates:
        service.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={'valueInputOption': 'RAW', 'data': updates}
        ).execute()
        print(f"  구독자 업데이트 완료 → 링크드인 팔로워: {linkedin_followers}, 유튜브 구독자: {youtube_subscribers}")


def main():
    print("Google Sheets 동기화 시작...")

    service = get_sheets_service()
    data = load_data_js()

    print("구독자/팔로워 수 업데이트 중...")
    update_subscriber_counts(service, data)

    existing = get_existing_data(service)
    print(f"시트 기존 항목: {len(existing)}개")

    def normalize(title):
        """특수문자·공백·이모지 제거, x/X/× 통일 후 한글+영문+숫자만 남김"""
        import unicodedata
        # × → x 통일
        title = title.replace('×', 'x').replace('X', 'x')
        result = ''
        for ch in title:
            cat = unicodedata.category(ch)
            if cat.startswith('L') or cat.startswith('N'):
                result += ch.lower()  # 대소문자 통일
        return result

    existing_keys = {}  # (channel, norm_title) → norm_title 길이
    for e in existing:
        norm = normalize(e['title'])
        key = (e['channel'], norm)
        existing_keys[key] = norm

    def is_duplicate(title, date, channel):
        norm = normalize(title)
        # 정확 매칭
        if (channel, norm) in existing_keys:
            return True
        # 앞부분 매칭: 시트 제목이 더 짧을 수 있으므로 겹치는 앞부분만 비교
        for (ch, ex_norm) in existing_keys:
            if ch != channel:
                continue
            min_len = min(len(norm), len(ex_norm))
            if min_len >= 8 and norm[:min_len] == ex_norm[:min_len]:
                return True
        return False

    new_rows = []
    new_rows_keys = set()  # 이번 배치 내 중복 방지

    # 유튜브 영상 (2026년만)
    for video in data.get('youtube_videos_list', []):
        if not video.get('date', '').startswith('2026'):
            continue
        title = video.get('title', '').strip()
        date = video.get('date', '')
        batch_key = ('유튜브', normalize(title))
        if is_duplicate(title, date, '유튜브') or batch_key in new_rows_keys:
            print(f"  [중복 건너뜀] {title[:40]} ({date})")
            continue
        new_rows_keys.add(batch_key)
        month = int(date.split('-')[1]) if date else 0
        new_rows.append({'date': date, 'row': [month, title, date, '유튜브', '']})

    # 블로그 게시글 (2026년만)
    for post in data.get('blog_posts_list', []):
        if not post.get('date', '').startswith('2026'):
            continue
        title = post.get('title', '').strip()
        date = post.get('date', '')
        batch_key = ('블로그', normalize(title))
        if is_duplicate(title, date, '블로그') or batch_key in new_rows_keys:
            print(f"  [중복 건너뜀] {title[:40]} ({date})")
            continue
        new_rows_keys.add(batch_key)
        month = int(date.split('-')[1]) if date else 0
        new_rows.append({'date': date, 'row': [month, title, date, '블로그', post.get('category', '')]})

    # LinkedIn 게시물 (2026년만)
    for post in data.get('linkedin_posts_list', []):
        if not post.get('date', '').startswith('2026'):
            continue
        title = post.get('content', '').strip()
        date = post.get('date', '')
        batch_key = ('링크드인', normalize(title))
        if is_duplicate(title, date, '링크드인') or batch_key in new_rows_keys:
            print(f"  [중복 건너뜀] {title[:40]} ({date})")
            continue
        new_rows_keys.add(batch_key)
        month = int(date.split('-')[1]) if date else 0
        new_rows.append({'date': date, 'row': [month, title, date, '링크드인', '']})

    sheet_id = get_sheet_id(service)

    if not new_rows:
        print("\n추가할 새 항목이 없습니다.")
        return

    new_rows.sort(key=lambda x: x['date'])

    print(f"\n새로 추가할 항목: {len(new_rows)}개")
    for item in new_rows:
        print(f"  [{item['row'][3]}] {item['row'][1][:40]} ({item['row'][2]})")

    import sys
    if sys.stdin.isatty():
        confirm = input("\n시트에 추가하시겠습니까? (y/n): ").strip().lower()
        if confirm != 'y':
            print("취소되었습니다.")
            return

    append_rows(service, [item['row'] for item in new_rows])
    print(f"완료! {len(new_rows)}개 항목이 추가되었습니다.")


if __name__ == '__main__':
    main()
