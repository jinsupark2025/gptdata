# Python으로 Microsoft Teams 채팅 송수신

이 저장소는 Microsoft Graph API를 사용해 Teams 채팅방에 메시지를 보내고, 최근 메시지를 읽고, 폴링으로 새 메시지를 확인하는 CLI 예제입니다.

## 1) 사전 준비

1. Azure Portal에서 App Registration 생성
2. Authentication에서 **Allow public client flows** 활성화
3. API permissions(Delegated) 추가
   - `Chat.Read`
   - `Chat.ReadWrite`
   - `ChatMessage.Send`
4. 필요 시 관리자 동의(Admin consent)

## 2) 설치

```bash
python -m venv .venv
source .venv/bin/activate
pip install msal requests
```

## 3) 환경 변수 설정

```bash
export TEAMS_TENANT_ID="<your-tenant-id>"
export TEAMS_CLIENT_ID="<your-client-id>"
```

## 4) 사용법

### 메시지 전송

```bash
python teams_chat_client.py send --chat-id "19:xxx@thread.v2" --message "안녕하세요 팀즈!"
```

### 최근 메시지 조회

```bash
python teams_chat_client.py read --chat-id "19:xxx@thread.v2" --top 10
```

### 실시간에 가까운 폴링 조회

```bash
python teams_chat_client.py watch --chat-id "19:xxx@thread.v2" --interval 5
```

최초 실행 시 디바이스 코드(Device Code) 로그인이 진행됩니다.
