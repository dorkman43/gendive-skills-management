# gendive-skills-management (경영 분석·보고)

경영 담당자와 팀장이 월간 실적을 분석하고 보고하는 업무에 쓰는 AI 스킬 모음입니다. 자동차 정비 서비스센터를 기준 예로 들었지만 일반 회사 경영 보고에도 쓸 수 있습니다. 매출 전년 동월 대비 분석, 센터별 생산성, 장기 미방문 고객 안내, 경영진 보고 요약, 실행계획 5가지를 다룹니다. (스킬: AI에게 "이 업무는 이런 순서와 기준으로 해 달라"고 미리 알려 두는 지침 파일)

## 스킬 목록

| 폴더명 | 한국어 이름 | 하는 일 | 이렇게 말하면 작동 |
|---|---|---|---|
| `sales-yoy-analysis` | 매출 추이·전년 동월 대비 분석 | 월별 추이, 구성비, 전년 동월 대비를 영업일수 보정(일평균)과 함께 보여주고 증감을 대수·객단가·차종 구성으로 나눕니다 | "전년 동월 대비 매출 분석해줘", "9월 매출이 왜 줄었는지 봐줘" |
| `utilization-analysis` | 센터별 생산성·가동률 | 가동률, 작업효율, 종합생산성, 베이 회전율, 대기 비중, 재수리율을 센터별로 계산하고 개선 포인트를 정리합니다 | "센터별 가동률 봐줘", "차가 오래 기다리는 센터가 어디야" |
| `lapsed-customer-outreach` | 장기 미방문 고객 안내 | 방문 간격으로 이탈 등급을 나누고 정기검사 도래 차량을 찾아, 수신동의와 광고 표시 규칙을 지킨 문자 초안을 만듭니다(발송하지 않음) | "오랫동안 안 온 고객 정리해줘", "정기검사 안내 문자 만들어줘" |
| `exec-summary-report` | 경영진 보고 요약 | 분석 결과를 1쪽 개조식·두괄식 보고로 줄입니다 | "경영진 보고용으로 한 장 요약해줘" |
| `action-plan` | 개선과제와 실행계획 | 과제 우선순위, 실행계획 표, 회의록에서 할 일 뽑기(담당·기한 누락 표시), 월간 점검표를 만듭니다 | "회의록에서 할 일만 뽑아줘", "실행계획 짜줘" |

### 같이 쓰면 좋은 공통 스킬 (`gendive-skills-common`)

| 상황 | 스킬 |
|---|---|
| 고객·직원 개인정보가 있는 자료를 AI에 올리기 전 | `data-masking` |
| 병합 셀, 소계 행, 월 표기가 제각각인 엑셀 | `excel-cleanup` |
| 보고하기 전 합계·증감률·기간 점검 | `report-qa` |
| 추이·비교 그래프, 대시보드 | `chart-dashboard` |

## 쓰는 법 3가지

1. **Claude 웹 (무료 포함)**
   - [Releases](https://github.com/dorkman43/gendive-skills-management/releases/latest)에서 쓸 스킬의 zip(예: `sales-yoy-analysis.zip`)을 내려받습니다. **압축을 풀지 말고 그대로** 올립니다. (Mac Safari가 자동으로 풀었다면 폴더를 오른쪽 클릭 → "압축"으로 다시 묶으세요)
   - claude.ai 설정 → Capabilities에서 "Code execution and file creation"(코드 실행 및 파일 생성)을 켭니다.
   - Customize → Skills → "+" → "Create skill" → "Upload a skill"에서 zip을 올리고, 목록의 스위치가 켜져 있는지 확인합니다.
   - 스킬은 사람마다 자기 계정에 올려야 합니다. 올린 뒤에는 파일을 첨부하고 평소처럼 말하면 알맞은 스킬이 자동으로 쓰입니다.
2. **Claude Code / Cowork**: `/plugin marketplace add dorkman43/gendive-skills-management` 를 입력한 뒤 `/plugin install gendive-skills-management@gendive-skills-management` 로 설치합니다.
3. **ChatGPT·Gemini**: `SKILL.md`의 맨 위 머리말(`---` 두 줄 사이)을 뺀 본문을 프로젝트 지침(ChatGPT) 또는 대화 첫 메시지에 붙여 넣고 파일을 올립니다. 각 서비스의 지침 글자 수 한도는 사용 전에 확인하세요.

## 실습용 가상 데이터

`samples/` 폴더에 연습용 파일이 있습니다. 센터·고객·차종은 모두 가상이며, 파일별 함정과 검산용 정답은 `samples/README.md`에 있습니다.

| 파일 | 연습할 스킬 |
|---|---|
| `정비매출.xlsx` | `sales-yoy-analysis` (`excel-cleanup`, `report-qa`, `chart-dashboard`) |
| `센터별_작업시간.xlsx` | `utilization-analysis` |
| `고객방문이력_가명.xlsx` | `lapsed-customer-outreach` |
| `회의록_월간경영회의.txt` | `action-plan` |

같은 파일을 다시 만들려면 저장소 폴더의 터미널(명령 입력창)에서 `python3 samples/make_samples.py` 를 실행합니다(openpyxl 필요).

## 꼭 지켜 주세요

- **실제 회사·고객 데이터를 무료·개인용 AI 계정에 올리지 마세요.** 개인용 계정은 입력 내용이 AI 학습에 쓰일 수 있습니다. 회사가 승인한 도구인지, 학습에 쓰이지 않는 기업용 계약인지 먼저 확인하세요.
- 고객 자료는 가명 처리(`data-masking`)한 뒤 올립니다. 주민등록번호·계좌번호·비밀번호는 어떤 AI에도 입력하지 않습니다(개인정보보호위원회 이용자 가이드).
- 고객 문자는 이 스킬이 **초안**만 만듭니다. 발송 전 수신동의, 광고 표시 등은 담당 부서가 확인하세요. 법률 자문이 아닙니다.
- 생산성 지표(가동률, 작업효율, 베이 회전율)는 국내 공식 정의를 확인하지 못해 해외 업계 방식과 이 스킬의 정의를 썼습니다. 회사 정의가 있으면 그것을 쓰세요.
- 판정 기준 숫자는 출발점입니다. 회사 기준이 있으면 바꿔 쓰세요. 회사 보안 규정이 이 저장소의 안내보다 우선합니다.
- 참고한 원본과 라이선스는 `THIRD_PARTY.md`를 보세요.

## 이용 조건

Copyright (c) 2026 dorkman43. 이 저장소는 PolyForm Internal Use License 1.0.0(`LICENSE`)에 개인 학습·실습 허락을 더한 조건을 따릅니다. 회사 내부 업무와 개인 학습에는 자유롭게 쓰고 고칠 수 있지만, 배포·공개나 강의·컨설팅 등 다른 사람에게 제공하는 데 쓰려면 허락이 필요합니다. 자세한 내용은 `NOTICE.md`를 보세요. 허락 문의: info@gendive.ai
