# Slack Channel Backup JSON to Excel Converter

슬랙 채널 백업 JSON 파일을 엑셀 파일로 변환하는 Python 도구입니다.

## 기능

- ✅ 슬랙 백업 JSON 파일을 읽어서 엑셀로 변환
- ✅ **폴더 내 여러 JSON 파일을 하나의 엑셀 파일로 통합**
- ✅ 사용자 정보, 채널 정보, 메시지 정보를 별도 시트로 정리
- ✅ 채널별 메시지를 개별 시트로 분리
- ✅ **날짜별 메시지를 개별 시트로 분리**
- ✅ **통계 정보 자동 생성** (날짜별, 사용자별, 파일별 메시지 수)
- ✅ 텍스트 정리 (HTML 태그, 이모지 제거)
- ✅ **줄바꿈 유지 및 엑셀에서 자동 줄바꿈 표시**
- ✅ 리액션, 파일, 첨부파일 정보 추출
- ✅ 타임스탬프를 읽기 쉬운 날짜/시간으로 변환

## 설치

1. 필요한 Python 패키지 설치:
```bash
pip install -r requirements.txt
```

## 사용법

### 1. 단일 JSON 파일 변환
```bash
python slack_to_excel.py your_slack_backup.json
```

### 2. 폴더 내 모든 JSON 파일 통합 변환 (새로운 기능!)
```bash
python slack_to_excel.py your_folder_path --folder
```

### 3. 출력 파일명 지정
```bash
python slack_to_excel.py your_slack_backup.json -o output.xlsx
python slack_to_excel.py your_folder_path --folder -o output.xlsx
```

## 출력 엑셀 파일 구조

### 단일 파일 변환 시
변환된 엑셀 파일에는 다음과 같은 시트들이 포함됩니다:

#### 1. Users 시트
- user_id: 사용자 ID
- username: 사용자명
- real_name: 실제 이름
- display_name: 표시 이름
- email: 이메일
- is_bot: 봇 여부
- is_admin: 관리자 여부
- is_owner: 소유자 여부
- deleted: 삭제된 사용자 여부
- created: 생성일
- updated: 수정일

#### 2. Channels 시트
- channel_id: 채널 ID
- channel_name: 채널명
- channel_type: 채널 타입 (public/private)
- topic: 채널 주제
- purpose: 채널 목적
- member_count: 멤버 수
- created: 생성일
- creator: 생성자

#### 3. Messages 시트
- message_id: 메시지 ID
- timestamp: 타임스탬프
- datetime: 날짜/시간
- user_id: 사용자 ID
- username: 사용자명
- text: 메시지 내용
- type: 메시지 타입
- subtype: 메시지 서브타입
- thread_ts: 스레드 타임스탬프
- reply_count: 답글 수
- reply_users_count: 답글 작성자 수
- latest_reply: 최신 답글
- reactions: 리액션 정보
- files: 첨부 파일 정보
- attachments: 첨부파일 정보

#### 4. 채널별 시트
각 채널의 메시지들이 개별 시트로 분리되어 저장됩니다.

### 폴더 통합 변환 시 (새로운 기능!)
폴더 내 모든 JSON 파일을 통합하여 다음과 같은 시트들이 생성됩니다:

#### 1. All_Messages 시트
- 모든 파일의 메시지가 통합되어 저장
- source_file: 원본 파일명
- 기타 메시지 정보 (위와 동일)

#### 2. Date_YYYY-MM-DD 시트들
- 각 날짜별로 메시지가 개별 시트로 분리
- 날짜순으로 정렬된 메시지들

#### 3. Statistics 시트
- **날짜별_메시지수**: 각 날짜의 메시지 수
- **사용자별_메시지수**: 각 사용자가 작성한 메시지 수
- **파일별_메시지수**: 각 JSON 파일의 메시지 수

## 예시

```bash
# 단일 슬랙 백업 파일 변환
python slack_to_excel.py slack_backup_2024.json

# 폴더 내 모든 JSON 파일 통합 변환
python slack_to_excel.py _note_scrum --folder

# 결과: _note_scrum_converted.xlsx 파일이 생성됩니다
# - 총 294개 파일, 373개 메시지가 통합됨
```

## 실제 사용 사례

`_note_scrum` 폴더의 경우:
- **입력**: 294개의 날짜별 JSON 파일 (2024-03-10 ~ 2025-01-15)
- **출력**: `_note_scrum_with_linebreaks.xlsx` (720KB)
- **내용**: 
  - All_Messages 시트: 모든 메시지 통합 (줄바꿈 유지)
  - Date_YYYY-MM-DD 시트: 날짜별 분리 (약 300개 시트, 줄바꿈 유지)
  - Statistics 시트: 통계 정보
- **특징**: 
  - 슬랙 메시지의 줄바꿈이 엑셀에서 그대로 표시됨
  - 스크럼 업데이트 형식 (`[이름]\n완료\nㄴ -`)이 보기 좋게 정리됨

## 주의사항

- JSON 파일은 UTF-8 인코딩이어야 합니다
- 대용량 파일의 경우 변환 시간이 오래 걸릴 수 있습니다
- Excel 시트명은 31자로 제한됩니다 (긴 채널명은 자동으로 축약됨)
- **폴더 통합 시 많은 시트가 생성될 수 있습니다** (날짜별 시트)

## 문제 해결

### 오류: "No module named 'pandas'"
```bash
pip install pandas openpyxl
```

### 오류: "File not found"
JSON 파일 경로가 올바른지 확인하세요.

### 오류: "Invalid JSON"
JSON 파일이 손상되었거나 올바른 형식이 아닙니다.

### 오류: "Too many worksheets"
Excel은 최대 1,048,576개의 시트를 지원합니다. 매우 많은 파일이 있는 경우 일부만 처리하거나 다른 방법을 고려하세요.
