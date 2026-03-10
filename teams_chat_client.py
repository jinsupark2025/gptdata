#!/usr/bin/env python3
"""Microsoft Teams chat sender/receiver via Microsoft Graph API.

Prerequisites
-------------
1) Azure App Registration (Public client)
2) Delegated API permissions (and admin consent when needed):
   - Chat.Read
   - Chat.ReadWrite
   - ChatMessage.Send
3) Environment variables:
   - TEAMS_TENANT_ID
   - TEAMS_CLIENT_ID

Usage examples
--------------
Send message:
    python teams_chat_client.py send --chat-id <CHAT_ID> --message "안녕하세요"

Read latest messages:
    python teams_chat_client.py read --chat-id <CHAT_ID> --top 10

Watch mode (polling every 5 seconds):
    python teams_chat_client.py watch --chat-id <CHAT_ID> --interval 5
"""

from __future__ import annotations

import argparse
import os
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from typing import Any

import msal
import requests

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
SCOPES = [
    "Chat.Read",
    "Chat.ReadWrite",
    "ChatMessage.Send",
]


class TeamsApiError(RuntimeError):
    """Raised when a Microsoft Graph API call fails."""


@dataclass
class TeamsConfig:
    tenant_id: str
    client_id: str

    @property
    def authority(self) -> str:
        return f"https://login.microsoftonline.com/{self.tenant_id}"


class TeamsChatClient:
    def __init__(self, config: TeamsConfig) -> None:
        self.config = config
        self.app = msal.PublicClientApplication(
            client_id=config.client_id,
            authority=config.authority,
        )
        self.session = requests.Session()
        self.session.headers["Content-Type"] = "application/json"

    def get_access_token(self) -> str:
        accounts = self.app.get_accounts()
        if accounts:
            token_result = self.app.acquire_token_silent(SCOPES, account=accounts[0])
            if token_result and "access_token" in token_result:
                return token_result["access_token"]

        flow = self.app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise TeamsApiError(f"Device flow initialization failed: {flow}")

        print(flow["message"])
        token_result = self.app.acquire_token_by_device_flow(flow)
        if "access_token" not in token_result:
            raise TeamsApiError(f"Authentication failed: {token_result}")

        return token_result["access_token"]

    def _request(self, method: str, endpoint: str, token: str, **kwargs: Any) -> dict[str, Any]:
        url = f"{GRAPH_BASE_URL}{endpoint}"
        headers = kwargs.pop("headers", {})
        headers["Authorization"] = f"Bearer {token}"

        response = self.session.request(method, url, headers=headers, timeout=30, **kwargs)
        if not response.ok:
            raise TeamsApiError(
                f"{method} {endpoint} failed ({response.status_code}): {response.text}"
            )

        if response.status_code == 204:
            return {}

        return response.json()

    def send_message(self, chat_id: str, message: str) -> dict[str, Any]:
        token = self.get_access_token()
        payload = {
            "body": {
                "contentType": "html",
                "content": message,
            }
        }
        return self._request("POST", f"/chats/{chat_id}/messages", token, json=payload)

    def read_messages(self, chat_id: str, top: int = 10) -> list[dict[str, Any]]:
        token = self.get_access_token()
        data = self._request(
            "GET",
            f"/chats/{chat_id}/messages?$top={top}",
            token,
        )
        return data.get("value", [])


def _parse_created_at(message: dict[str, Any]) -> datetime:
    created = message.get("createdDateTime")
    if not created:
        return datetime.min
    return datetime.fromisoformat(created.replace("Z", "+00:00"))


def print_messages(messages: list[dict[str, Any]]) -> None:
    if not messages:
        print("메시지가 없습니다.")
        return

    ordered = sorted(messages, key=_parse_created_at)
    for msg in ordered:
        sender = (
            msg.get("from", {})
            .get("user", {})
            .get("displayName", "Unknown")
        )
        created = msg.get("createdDateTime", "")
        content = msg.get("body", {}).get("content", "")
        print(f"[{created}] {sender}: {content}")


def require_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise ValueError(f"환경 변수 {name} 이(가) 필요합니다.")
    return value


def build_client() -> TeamsChatClient:
    config = TeamsConfig(
        tenant_id=require_env("TEAMS_TENANT_ID"),
        client_id=require_env("TEAMS_CLIENT_ID"),
    )
    return TeamsChatClient(config)


def main() -> int:
    parser = argparse.ArgumentParser(description="Teams 채팅 송수신 CLI")
    subparsers = parser.add_subparsers(dest="command", required=True)

    send_parser = subparsers.add_parser("send", help="메시지 전송")
    send_parser.add_argument("--chat-id", required=True, help="Teams chat ID")
    send_parser.add_argument("--message", required=True, help="전송할 메시지")

    read_parser = subparsers.add_parser("read", help="메시지 조회")
    read_parser.add_argument("--chat-id", required=True, help="Teams chat ID")
    read_parser.add_argument("--top", type=int, default=10, help="가져올 메시지 수")

    watch_parser = subparsers.add_parser("watch", help="메시지 폴링")
    watch_parser.add_argument("--chat-id", required=True, help="Teams chat ID")
    watch_parser.add_argument("--interval", type=int, default=5, help="폴링 주기(초)")

    args = parser.parse_args()

    try:
        client = build_client()

        if args.command == "send":
            sent = client.send_message(args.chat_id, args.message)
            print(f"전송 완료: message id={sent.get('id')}")
            return 0

        if args.command == "read":
            messages = client.read_messages(args.chat_id, args.top)
            print_messages(messages)
            return 0

        if args.command == "watch":
            seen_ids: set[str] = set()
            print("watch 모드 시작 (Ctrl+C 종료)")
            while True:
                messages = client.read_messages(args.chat_id, 20)
                new_messages = [m for m in messages if m.get("id") not in seen_ids]
                if new_messages:
                    print_messages(new_messages)
                    for message in new_messages:
                        message_id = message.get("id")
                        if message_id:
                            seen_ids.add(message_id)
                time.sleep(args.interval)

    except KeyboardInterrupt:
        print("\n종료합니다.")
        return 0
    except (ValueError, TeamsApiError, requests.RequestException) as exc:
        print(f"오류: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
