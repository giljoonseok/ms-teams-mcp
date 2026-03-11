# "어제 메일 요약해줘" 한마디로 끝 — AI 업무 자동화 도구 안내

> AI 비서(Claude)에게 말만 하면 Teams 메시지, Outlook 이메일, 캘린더, 파일을 알아서 처리해줍니다.

---

## 이런 게 가능합니다

| 기존 방식 | AI 활용 |
|-----------|---------|
| Teams 앱 열어서 메시지 하나하나 확인 | **"안 읽은 메시지 요약해줘"** 한마디로 끝 |
| Outlook에서 메일 검색하고 답장 작성 | **"김팀장 메일 읽고 답장해줘"** |
| 캘린더 열어서 빈 시간 찾고 일정 생성 | **"내일 오후 2시 회의 잡아줘"** |
| 여러 앱 오가며 정보 취합 | **"메일, 채팅, 일정 종합해서 브리핑 해줘"** |

---

## 목차

1. [설치하기](#1-설치하기)
2. [Microsoft 계정 연결 (인증)](#2-microsoft-계정-연결-인증)
3. [이렇게 사용하세요 — 업무 활용 예제](#3-이렇게-사용하세요--업무-활용-예제)
4. [할 수 있는 일 전체 목록](#4-할-수-있는-일-전체-목록)
5. [문제가 생겼을 때](#5-문제가-생겼을-때)

---

## 1. 설치하기

> 총 4단계입니다. 약 10분 소요됩니다.

### 1단계: Python 설치 (필수)

> 이미 Python이 설치되어 있다면 이 단계를 건너뛰세요. PowerShell에서 `python --version` 입력 시 3.10 이상이면 OK.

**Windows:**
1. [https://www.python.org/downloads/](https://www.python.org/downloads/) 에서 최신 버전 다운로드
2. 설치 시 **"Add python.exe to PATH"** 체크박스를 반드시 선택
3. 설치 완료 후 PowerShell을 닫았다가 다시 열기

**Mac:** (기본 설치되어 있음. 없으면 `brew install python`)

### 2단계: uv 설치 (도구 실행기)

**Windows** — PowerShell을 열고 아래 명령어를 복사/붙여넣기:
```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

**Mac:**
```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

설치 후 **터미널(PowerShell)을 닫았다가 다시 열어주세요.**

### 3단계: Claude Desktop 설치

1. [https://claude.com/download](https://claude.com/download) 에서 다운로드
2. 설치 파일 실행
3. Anthropic 계정으로 로그인

### 4단계: Teams 연동 설정

Claude Desktop의 설정 파일을 열어 아래 내용을 붙여넣습니다.

**설정 파일 여는 방법:**
- Claude Desktop 실행 > 좌측 하단 **Settings** > **Developer** 탭 > **Edit Config** 클릭

또는 직접 파일 열기:
- **Windows**: 파일 탐색기 주소창에 `%APPDATA%\Claude` 입력 > `claude_desktop_config.json` 메모장으로 열기
- **Mac**: `~/Library/Application Support/Claude/claude_desktop_config.json`

**아래 내용을 파일에 붙여넣기** (기존 내용이 있으면 전체 교체):

```json
{
  "mcpServers": {
    "ms-teams": {
      "command": "uvx",
      "args": [
        "--from",
        "ms-teams-mcp",
        "ms-teams-mcp"
      ],
      "env": {
        "MS_CLIENT_ID": "<YOUR_CLIENT_ID>",
        "MS_CLIENT_SECRET": "<YOUR_CLIENT_SECRET>",
        "MS_TENANT_ID": "<YOUR_TENANT_ID>"   
      }
    }
  }
}
```

> **주의: 위 키 값은 사내 전용입니다. 외부(블로그, GitHub, SNS 등)에 절대 공유하지 마세요.**

> **Windows에서 `uvx`가 인식되지 않는 경우**: `"command": "uvx"` 를 `"command": "uvx.cmd"` 로 변경

파일 저장 후 Claude Desktop을 **완전히 종료(트레이 아이콘 우클릭 > 종료) 후 다시 실행**합니다.

---

## 2. Microsoft 계정 연결 (인증)

처음 사용할 때 1회만 하면 됩니다.

Claude Desktop 대화창에 아래와 같이 입력:

```
팀즈 인증해줘
```

Claude가 안내하는 대로 따라하세요:

1. Claude가 보여주는 **링크(https://microsoft.com/devicelogin)** 를 브라우저에서 열기
2. Claude가 보여주는 **코드** 입력
3. 회사 Microsoft 계정으로 **로그인**
4. 권한 **동의**

인증 완료! 이후에는 자동 로그인됩니다.

---

## 3. 이렇게 사용하세요 — 업무 활용 예제

Claude Desktop 대화창에 자연스러운 한국어로 말하면 됩니다.

### 아침 출근 — 하루 시작 브리핑

```
안 읽은 메시지랑 메일 확인해서 오늘 아침 브리핑 해줘
오늘 일정 뭐 있어?
```

### Teams 채팅

```
최근 채팅 목록 보여줘
유영수 이사님과의 채팅 내용 보여줘
김팀장에게 "회의 자료 검토 부탁드립니다" 메시지 보내줘
```

### Teams 채널

```
내 팀 목록 보여줘
"개발팀" 채널 최근 메시지 보여줘
#general 채널에 "오늘 점심 회식입니다" 메시지 보내줘
```

### 이메일

```
오늘 받은 메일 확인해줘
안 읽은 메일 보여줘
김과장님한테 온 최신 메일 읽어줘
"프로젝트 보고서" 관련 메일 검색해줘
김과장님 메일에 "검토 완료, 진행해주세요" 답장해줘
홍길동@example.com에게 "주간 보고" 제목으로 이번 주 진행 사항 정리해서 메일 보내줘
```

### 캘린더 / 일정

```
이번 주 일정 보여줘
내일 회의 뭐 있어?
내일 오후 2시에 홍길동@example.com과 1시간 "디자인 리뷰" 회의 잡아줘
다음 주 월요일부터 매일 오전 9시 30분 스탠드업 반복 일정 만들어줘
오후 2시 회의를 4시로 변경해줘
금요일 회의 취소해줘
내일 오후 3시에 보고서 마감 리마인더 설정해줘
```

### 파일 검색

```
"프로젝트A" 팀 채널의 파일 목록 보여줘
project-plan.xlsx 파일 읽어줘
파일에서 "예산" 검색해줘
```

### 사용자 검색

```
조직에서 "김" 이름 가진 사용자 검색해줘
마케팅팀 담당자 이메일 주소 찾아줘
```

### 복합 업무 — 이게 진짜 강력합니다

```
어제 팀즈, 이메일 내용 요약해줘
```
```
이번 주 받은 메일 중 "프로젝트X" 관련 내용 전부 찾아서 요약하고,
팀 채널에 주간 현황 공유해줘
```
```
다음 주 캘린더 확인하고 1시간짜리 회의 가능한 시간 찾아줘
```
```
메일이랑 파일에서 "예산" 검색해서 주요 수치 정리해줘
```

---

## 4. 할 수 있는 일 전체 목록

| 분류 | 기능 |
|------|------|
| **Teams 채팅** | 채팅 목록 보기, 메시지 읽기, 메시지 보내기, 답장, 새 채팅 만들기 |
| **Teams 채널** | 팀/채널 목록, 채널 메시지 읽기, 메시지 보내기, 답장 |
| **이메일** | 메일 목록, 메일 읽기, 메일 검색, 메일 보내기, 답장, 전달, 폴더 목록 |
| **캘린더** | 일정 조회, 일정 생성/수정/삭제, 반복 일정, 리마인더 |
| **파일** | 채널 파일 목록, 파일 읽기(텍스트/엑셀), 파일 인덱스 구축, 파일 검색 |
| **사용자** | 조직 내 사용자 검색 |
| **요약** | 안 읽은 메일/채팅 종합 요약 |
| **시스템** | 인증 상태 확인, 버전 업데이트 확인 및 자동 업데이트 |

---

## 5. 문제가 생겼을 때

| 증상 | 해결 방법 |
|------|-----------|
| Claude에서 Teams 기능이 안 보임 | Claude Desktop을 완전히 종료(트레이 아이콘까지) 후 재시작 |
| "인증이 필요합니다" 메시지 | `팀즈 인증해줘` 입력하여 재인증 |
| Windows에서 서버가 안 뜸 | 설정 파일의 `"command": "uvx"` 를 `"command": "uvx.cmd"` 로 변경 |
| 권한 오류 | IT 관리자에게 Azure 앱 권한 확인 요청 |
| 갑자기 동작이 안 됨 | Claude에게 `MCP 서버 업데이트 확인해줘` 입력 |

### 업데이트

서버 시작 시 **자동으로 최신 버전을 확인하고 업데이트**합니다. (1일 1회 체크)

수동으로 업데이트하려면 PowerShell에서:
```powershell
pip install --upgrade ms-teams-mcp
```

현재 설치된 버전 확인:
```powershell
ms-teams-mcp --version
```

uvx 캐시 초기화 (uvx 사용 시):
```powershell
uv cache clean
```

이후 Claude Desktop 재시작하면 최신 버전이 적용됩니다.

---

*문의: 길준석 (jun@3hs.co.kr)*
