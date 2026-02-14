# SharePoint MCP Server — Product Requirements Document

> **Remote Model Context Protocol Server for Microsoft SharePoint Integration**

| 항목 | 내용 |
|------|------|
| **Project Name** | SharePoint MCP Server |
| **Version** | 1.0 |
| **Author** | Logan |
| **Date** | 2026-02-14 |
| **Status** | Draft |
| **Deployment** | Cloudflare Workers (무료) |
| **Scope** | PoC / 개인 학습용 |

---

## 1. Executive Summary

SharePoint MCP Server는 Microsoft SharePoint Online에 저장된 문서, 파일, 리스트 데이터를 AI 클라이언트(Claude, ChatGPT, Cursor 등)에서 직접 접근할 수 있도록 하는 Remote MCP(Model Context Protocol) 서버입니다.

본 프로젝트는 PoC(Proof of Concept) 및 개인 학습 목적으로, Cloudflare Workers 무료 플랜을 활용하여 인프라 비용 없이 배포합니다. Microsoft Graph API를 통해 SharePoint와 통신하며, MCP 표준 프로토콜을 준수하여 다양한 AI 플랫폼에서 범용적으로 사용 가능합니다.

---

## 2. Problem Statement

### 2.1 현재 문제점

- SharePoint에 저장된 문서와 데이터를 AI와 함께 활용하려면 수동으로 다운로드/복사 후 붙여넣기 필요
- SharePoint 리스트 데이터를 조회하거나 수정할 때마다 브라우저와 AI 툴 사이를 번갈아 전환해야 함
- AI 툴별로 별도의 연동 개발이 필요하여 유지보수 부담 증가

### 2.2 목표 상태

- AI 대화창에서 자연어로 지시하면 SharePoint 문서 검색/조회 가능
- SharePoint 리스트 데이터를 AI를 통해 조회/생성/수정/삭제 가능
- 하나의 MCP 서버로 Claude, ChatGPT, Cursor 등 다양한 AI 클라이언트에서 사용

---

## 3. Scope

### 3.1 In Scope (Phase 1 — PoC)

| 기능 영역 | 세부 기능 | 우선순위 |
|-----------|----------|---------|
| 문서/파일 검색 | 키워드 기반 문서 검색, 파일 메타데이터 조회 | P0 (Must-Have) |
| 문서 다운로드 | 파일 콘텐츠 다운로드 및 텍스트 추출 | P0 (Must-Have) |
| 리스트 조회 | SharePoint 리스트 아이템 조회 및 필터링 | P0 (Must-Have) |
| 리스트 생성 | 새 리스트 아이템 추가 | P1 (Should-Have) |
| 리스트 수정 | 기존 리스트 아이템 업데이트 | P1 (Should-Have) |
| 리스트 삭제 | 리스트 아이템 삭제 | P2 (Nice-to-Have) |

### 3.2 Out of Scope (Phase 1)

- SharePoint 사이트/페이지 관리 (CMS 기능)
- 권한 및 접근 관리 (Permission Management)
- 문서 업로드/생성 기능
- SharePoint On-Premise 지원
- 대용량 파일 처리 및 배치 작업

---

## 4. Technical Architecture

### 4.1 시스템 구성도

```
┌──────────────────┐     ┌─────────────────────┐     ┌──────────────────┐     ┌──────────────────┐
│   AI Client      │     │  Cloudflare Workers  │     │  Microsoft Graph │     │   SharePoint     │
│ (Claude/ChatGPT/ │────▶│   (MCP Server)       │────▶│     API v1.0     │────▶│    Online        │
│  Cursor)         │◀────│                      │◀────│                  │◀────│                  │
└──────────────────┘     └─────────────────────┘     └──────────────────┘     └──────────────────┘
                          MCP Protocol                 OAuth 2.0 Bearer        REST API
                          (Streamable HTTP)             Token
```

### 4.2 기술 스택

| 구성 요소 | 기술 | 선정 사유 |
|----------|------|----------|
| Runtime | Cloudflare Workers | 무료 플랜, 글로벌 배포, 콜드스타트 없음 |
| Language | TypeScript | MCP SDK 공식 지원, 타입 안정성 |
| MCP SDK | @modelcontextprotocol/sdk | Anthropic 공식 MCP SDK |
| Transport | Streamable HTTP | MCP 최신 표준, SSE 대체 |
| SharePoint API | Microsoft Graph API v1.0 | SharePoint REST API 대체, 통합 엔드포인트 |
| 인증 | Azure AD (Entra ID) OAuth 2.0 | Microsoft 공식 인증 방식 |
| 세션 저장소 | Cloudflare KV | OAuth 토큰 캐싱, 무료 플랜 포함 |

### 4.3 인증 플로우

```
1. AI 클라이언트가 MCP 서버에 접속 요청
2. MCP 서버가 Azure AD OAuth 2.0 인증 플로우 시작
3. 사용자가 Microsoft 계정으로 로그인 및 권한 동의
4. Access Token 발급 후 Cloudflare KV에 캐싱
5. 이후 요청에서 캐싱된 토큰으로 Microsoft Graph API 호출
```

### 4.4 프로젝트 구조

```
sharepoint-mcp-server/
├── src/
│   ├── index.ts              # 엔트리포인트, HTTP 라우팅
│   ├── server.ts             # MCP 서버 인스턴스 및 Tool 등록
│   ├── auth.ts               # Azure AD OAuth 처리
│   ├── graph-client.ts       # Microsoft Graph API 클라이언트
│   └── tools/
│       ├── documents.ts      # 문서 검색/조회/다운로드 Tools
│       └── lists.ts          # 리스트 CRUD Tools
├── wrangler.jsonc            # Cloudflare Workers 설정
├── package.json
└── tsconfig.json
```

---

## 5. MCP Tool Specification

### 5.1 문서/파일 관련 Tools

#### Tool: `search_documents`

**설명**: SharePoint 사이트 내 문서를 키워드로 검색합니다.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `query` | string | Yes | 검색 키워드 |
| `site_id` | string | No | 특정 사이트 ID (기본: 루트 사이트) |
| `file_type` | string | No | 파일 확장자 필터 (docx, xlsx, pdf 등) |
| `max_results` | number | No | 최대 결과 수 (기본: 10) |

**Response**: 파일 이름, 경로, 수정일, 크기, 작성자 목록

---

#### Tool: `get_file_content`

**설명**: 특정 파일의 콘텐츠를 다운로드하여 텍스트로 반환합니다.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `file_id` | string | Yes | 파일의 고유 ID (driveItem ID) |
| `site_id` | string | No | 사이트 ID |

**Response**: 파일 텍스트 콘텐츠 (지원 형식: txt, csv, json 등)

---

#### Tool: `list_files`

**설명**: 특정 폴더의 파일 목록을 조회합니다.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `folder_path` | string | No | 폴더 경로 (기본: 루트) |
| `site_id` | string | No | 사이트 ID |

**Response**: 파일/폴더 이름, 타입, 크기, 수정일 목록

---

### 5.2 리스트 관련 Tools

#### Tool: `get_list_items`

**설명**: SharePoint 리스트의 아이템들을 조회합니다.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `list_name` | string | Yes | 리스트 이름 또는 ID |
| `site_id` | string | No | 사이트 ID |
| `filter` | string | No | OData 필터 쿼리 |
| `select` | string[] | No | 반환할 필드 목록 |
| `top` | number | No | 최대 반환 건수 (기본: 50) |

**Response**: 리스트 아이템 배열 (필드 값 포함)

---

#### Tool: `create_list_item`

**설명**: SharePoint 리스트에 새 아이템을 추가합니다.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `list_name` | string | Yes | 리스트 이름 또는 ID |
| `fields` | object | Yes | 필드 값 객체 (key-value) |
| `site_id` | string | No | 사이트 ID |

**Response**: 생성된 아이템 ID 및 필드 값

---

#### Tool: `update_list_item`

**설명**: 기존 리스트 아이템을 수정합니다.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `list_name` | string | Yes | 리스트 이름 또는 ID |
| `item_id` | string | Yes | 아이템 ID |
| `fields` | object | Yes | 수정할 필드 값 |
| `site_id` | string | No | 사이트 ID |

**Response**: 수정된 아이템 필드 값

---

#### Tool: `delete_list_item`

**설명**: 리스트 아이템을 삭제합니다.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `list_name` | string | Yes | 리스트 이름 또는 ID |
| `item_id` | string | Yes | 삭제할 아이템 ID |
| `site_id` | string | No | 사이트 ID |

**Response**: 삭제 성공 여부

---

## 6. Prerequisites

### 6.1 Azure AD (Entra ID) 앱 등록

1. Azure Portal → Microsoft Entra ID → App registrations → New registration
2. Redirect URI 설정: `https://<worker-subdomain>.workers.dev/auth/callback`
3. API Permissions 추가:
   - `Sites.Read.All` — 사이트 및 문서 읽기
   - `Sites.ReadWrite.All` — 리스트 CRUD (쓰기 필요 시)
   - `Files.Read.All` — 파일 콘텐츠 읽기
4. Client Secret 생성 및 보관

### 6.2 Cloudflare 계정 설정

1. Cloudflare 무료 계정 생성
2. Workers KV Namespace 생성 (세션/토큰 저장용)
3. Wrangler CLI 설치 및 로그인

### 6.3 환경 변수 (시크릿)

| 변수명 | 설명 |
|--------|------|
| `AZURE_CLIENT_ID` | Azure AD 앱 클라이언트 ID |
| `AZURE_CLIENT_SECRET` | Azure AD 앱 클라이언트 시크릿 |
| `AZURE_TENANT_ID` | Azure AD 테넌트 ID |
| `SHAREPOINT_SITE_URL` | 기본 SharePoint 사이트 URL |

---

## 7. Microsoft Graph API Endpoints

본 프로젝트에서 사용하는 주요 Graph API 엔드포인트입니다.

| 기능 | Method | Endpoint |
|------|--------|----------|
| 문서 검색 | GET | `/sites/{site-id}/drive/root/search(q='{query}')` |
| 파일 목록 | GET | `/sites/{site-id}/drive/root:/{path}:/children` |
| 파일 다운로드 | GET | `/sites/{site-id}/drive/items/{item-id}/content` |
| 리스트 조회 | GET | `/sites/{site-id}/lists/{list-id}/items` |
| 리스트 생성 | POST | `/sites/{site-id}/lists/{list-id}/items` |
| 리스트 수정 | PATCH | `/sites/{site-id}/lists/{list-id}/items/{item-id}` |
| 리스트 삭제 | DELETE | `/sites/{site-id}/lists/{list-id}/items/{item-id}` |

---

## 8. Milestones & Timeline

PoC 프로젝트로 총 3주 기간을 목표로 합니다.

| Week | Milestone | 세부 작업 | 산출물 |
|------|-----------|----------|--------|
| Week 1 | 환경 설정 & 기본 구조 | Azure AD 앱 등록, CF Workers 설정, MCP 기본 서버, OAuth 플로우 | 인증 성공 |
| Week 2 | Tool 구현 | 문서 검색/조회 Tool, 리스트 CRUD Tool, 에러 핸들링 | Tool 동작 |
| Week 3 | 통합 테스트 & 문서화 | Claude/ChatGPT 연동 테스트, 버그 수정, README 작성 | PoC 완료 |

---

## 9. Success Criteria

| # | 기준 | 측정 방법 |
|---|------|----------|
| 1 | Claude에서 자연어 명령으로 SharePoint 문서 검색 성공 | 수동 테스트 |
| 2 | SharePoint 리스트 아이템 CRUD 작업 성공 | 수동 테스트 |
| 3 | OAuth 인증 플로우 정상 동작 | 인증 플로우 테스트 |
| 4 | Cloudflare Workers 무료 플랜 범위 내 운영 | CF Dashboard 모니터링 |
| 5 | Claude 및 ChatGPT 양쪽에서 연동 확인 | 수동 테스트 |

---

## 10. Risks & Mitigation

| 리스크 | 영향도 | 완화 방안 |
|--------|--------|----------|
| Azure AD 권한 부족 | 높음 | 사전에 필요 권한 목록 확인, Admin Consent 필요 시 IT 팀 협조 |
| CF Workers 무료 플랜 제약 | 낮음 | PoC 범위에서는 충분, 확장 시 $5/월 유료 플랜 검토 |
| Graph API Rate Limiting | 중간 | 요청 캐싱 및 배치 처리 고려, Retry with backoff 구현 |
| MCP 프로토콜 변경 | 낮음 | MCP SDK 버전 고정, Breaking change 모니터링 |
| 토큰 보안 이슈 | 높음 | KV 암호화, 토큰 TTL 설정, HTTPS 필수 |

---

## 11. Future Considerations (Phase 2+)

- 문서 업로드 기능 추가 (AI가 생성한 문서를 SharePoint에 저장)
- SharePoint 사이트/페이지 관리 기능
- Webhook 연동을 통한 실시간 알림 (MCP Notifications)
- SharePoint On-Premise 지원 (하이브리드 환경)
- Multi-tenant 지원으로 조직 전체 배포
- Azure Functions로 마이그레이션 (사내 인프라 통합 시)
- MCP Resource/Prompt 타입 추가 (Tool 외 MCP 기능 활용)

---

## 12. References

- [MCP Specification](https://spec.modelcontextprotocol.io)
- [MCP TypeScript SDK](https://github.com/modelcontextprotocol/typescript-sdk)
- [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/api/overview)
- [Cloudflare Workers](https://developers.cloudflare.com/workers)
- [Cloudflare MCP Template](https://github.com/cloudflare/ai/tree/main/demos/remote-mcp-server)