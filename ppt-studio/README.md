# PPT 스튜디오

PPT 제작 전용 프로젝트. **세팅을 먼저 고정하고(디자인 시스템·지침·수정 기록), 단계별로 만든다**는 원칙으로 구성함.

## 바로 쓰는 법
| 상황 | 할 일 |
|---|---|
| claude.ai에서 쓰기 | `claude-project/프로젝트지침.md` 보고 프로젝트 생성 → 지침 붙여넣기 → 지식 파일 업로드 |
| Claude Code(이 저장소)에서 쓰기 | "PPT 만들어줘" + 자료 첨부 → `ppt-studio` 스킬과 `CLAUDE.md`가 자동 적용 |
| 어떤 프롬프트를 쓸지 | `prompts/프롬프트팩.md` 0단계부터 |
| 왜 이렇게 하는지 | `docs/01_영상분석.md` |

## 구성
```
ppt-studio/
├─ CLAUDE.md                     작업 지침 (Claude Code용)
├─ claude-project/프로젝트지침.md  claude.ai 프로젝트 세팅·지침 원문
├─ docs/01_영상분석.md            영상 분석 + 적용 원칙 + 출처
├─ design-system/
│  ├─ design-system.md           고정 디자인 규칙
│  └─ brands/*.json              법인별 색·폰트·푸터 (CI 확인 전 임시값)
├─ templates/덱유형별_목차.md      IR·정부사업·강의·투심·제안서·내부보고
├─ prompts/프롬프트팩.md           단계별 프롬프트
├─ lessons-log.md                수정 기록 누적
├─ tools/build_deck.py           spec.json → 편집 가능한 .pptx (네이티브 차트·표)
├─ tools/render.sh               .pptx → PDF·PNG (검수용)
├─ examples/sample-ir.json       전 레이아웃 예시 (가상 데이터)
└─ projects/                     덱별 작업 폴더
```
스킬: 저장소 루트 `.claude/skills/ppt-studio/SKILL.md`

## 남은 확인 사항
- [ ] 법인별 로고 파일, CI 색상(HEX) — 받기 전까지 4개 법인 모두 중립 네이비·오렌지 임시값
- [ ] 커리즈 법인 형태 표기
- [ ] 영상 ① `byShnblO4Oc` 내용, 영상 ② 자막·댓글 원문 — 작업 환경에서 YouTube 접속 차단으로 미확보
- [ ] 레퍼런스로 삼을 잘 만든 덱 2~3개
