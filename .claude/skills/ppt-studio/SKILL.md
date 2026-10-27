---
name: ppt-studio
description: PPT·슬라이드·발표자료·IR 덱·정부사업 발표평가 자료·강의안·제안서를 만들거나 기존 PPT를 리디자인할 때 사용. 브리프 → 스토리라인 → 1장 테스트 → 10장 단위 빌드 → 렌더링 검수 순서로, 법인별 브랜드 토큰과 Pretendard 디자인 시스템을 적용한 편집 가능한 .pptx(네이티브 차트·표)와 PDF를 만든다.
---

# ppt-studio 스킬

먼저 `ppt-studio/CLAUDE.md`를 읽고 그 작업 순서와 절대 규칙을 따른다. 이 문서는 빌더 사용법이다.

## 파일 위치
| 경로 | 내용 |
|---|---|
| `ppt-studio/design-system/design-system.md` | 고정 디자인 규칙 |
| `ppt-studio/design-system/brands/*.json` | 법인별 색·폰트·푸터 토큰 |
| `ppt-studio/lessons-log.md` | 누적 수정 규칙 (작업 전 반드시 확인) |
| `ppt-studio/templates/덱유형별_목차.md` | 덱 유형별 목차 |
| `ppt-studio/prompts/프롬프트팩.md` | 단계별 프롬프트 |
| `ppt-studio/tools/build_deck.py` | spec.json → .pptx |
| `ppt-studio/tools/render.sh` | .pptx → PDF + 슬라이드 PNG |
| `ppt-studio/examples/sample-ir.json` | 전 유형 예시 spec |
| `ppt-studio/projects/<YYMMDD>_<이름>/` | 덱별 작업 폴더 (spec.json, 원자료, 산출물) |

## 실행
```bash
python3 ppt-studio/tools/build_deck.py ppt-studio/projects/<폴더>/spec.json ppt-studio/projects/<폴더>/<이름>.pptx
bash ppt-studio/tools/render.sh ppt-studio/projects/<폴더>/<이름>.pptx
```
- 렌더링에 Pretendard가 필요하다. 없으면 `npm pack pretendard@1.3.9`로 받아 `dist/public/static/*.otf`를 `~/.fonts`에 복사 후 `fc-cache -f`.
- 렌더링된 PNG를 **반드시 직접 열어 보고** 넘침·겹침·잘림을 확인한다. 빌더 경고와 `[확인필요]` 목록은 사용자 보고에 그대로 포함한다.

## spec.json 형식
```json
{
  "brand": "ztop | mademate | zerotoan | careez | default",
  "size": "16:9 | A4-landscape | A4-portrait",
  "slides": [ { "type": "...", ... } ]
}
```
공통 선택 키(본문 슬라이드): `chapter`(챕터명), `title`(결론 한 문장 40자 이내), `subtitle`, `source`(출처, 푸터에 표시), `notes`(발표자 노트)

| type | 필수 키 | 선택 키 | 용도 |
|---|---|---|---|
| `cover` | `title` | `subtitle`, `date`, `presenter` | 표지 |
| `agenda` | `items`[] | `title` | 목차 |
| `section` | `title` | `no` | 챕터 간지 |
| `bullets` | `bullets`[] (≤3) | `message`(강조 한 줄) | 주장 + 근거 |
| `kpi` | `items`[{`value`,`label`,`desc`}] (≤4) | | 핵심 숫자 |
| `chart` | `chart`(bar/stacked/hbar/line/pie), `categories`[], `series`[{`name`,`values`[]}] | `unit`, `takeaway`, `highlight`(강조 막대 인덱스), `number_format` | 네이티브 차트 |
| `table` | `columns`[], `rows`[[]] (≤8행) | `highlight_col` | 비교표 |
| `compare` | `left`/`right` {`title`,`bullets`[]} | | 전후·A/B 비교 |
| `process` | `steps`[{`title`,`desc`}] (≤5) | | 절차·일정·자금 사용 |
| `closing` | | `title`, `contact` | 마무리 |

- `pie`는 도넛 차트로 그리고 값은 비율(0.6 = 60%)로 넣는다.
- 텍스트에 `[확인필요]`를 넣으면 빨간색으로 표시되고 빌드 후 목록이 출력된다.
- 줄바꿈은 `\n`.

## 사용자 보고 형식
1. 결론: 파일 경로, 장수, 규격, 법인
2. `[확인필요]` 목록 (장 번호별)
3. 빌더 경고 중 조치하지 않은 것과 사유
4. 다음 액션 (사용자가 확인·결정할 것)
