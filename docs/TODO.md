# kotlin-excel-dsl TODO

## 현재 상태

**v1.0.0 핵심 기능 구현 완료**

---

## 완료된 기능 (v1.0.0)

### Phase 1: 핵심 기능 (MVP)
- [x] DSL 기본 동작 (excel/sheet/column/rows)
- [x] 어노테이션 기반 생성 (@Excel, @Column)
- [x] 기본 렌더링 (헤더 + 데이터)
- [x] 컬럼 너비 (auto, 고정 - `20.chars`)
- [x] 멀티 시트 지원

### Phase 2: 스타일링
- [x] CellStyle → POI 변환 (색상, 폰트, 테두리, 정렬)
- [x] 테마 적용 (`excel(theme = Theme.Modern)`)
- [x] CSS-like 스타일 DSL (`styles { header { ... } }`)
- [x] 미리 정의된 테마 (Modern, Minimal, Classic)
- [x] 인라인 컬럼 스타일 (`column("가격", style = { bold() })`)
- [x] 컬럼별 스타일 DSL (`styles { column("금액") { body { align(RIGHT) } } }`)

### Phase 3: 고급 기능
- [x] 헤더 그룹 (`headerGroup("제목") { ... }`)
- [x] 셀 병합 (헤더 그룹 자동 병합)
- [x] 날짜/시간 포맷 자동 감지 (LocalDate, LocalDateTime)

### Phase 4: 품질
- [x] E2E 테스트 (Kotest)
- [x] 컴파일 경고 정리
- [x] README.md / README.ko.md 작성
- [x] 상세한 에러 메시지 (힌트 포함)
- [x] 통합 모듈 (excel-dsl)

---

## 향후 버전 (v1.1+)

### 스타일링 확장
- [ ] 줄무늬 효과 (alternateRowStyle)
- [ ] 조건부 스타일 (`conditionalStyle { if (value < 0) fontColor(RED) }`)

### 어노테이션 확장
- [ ] `excelOf(data, theme = Theme.Modern)` 테마 파라미터
- [ ] `@HeaderStyle`, `@BodyStyle` 어노테이션

### 기능 추가
- [ ] 틀 고정 (freezePane)
- [ ] 자동 필터
- [ ] 조건부 서식 (POI native)
- [ ] 수식 지원
- [ ] Percent 너비 구현 (`30.percent`)

### 품질 개선
- [ ] 스타일 캐싱 (POI 스타일 개수 제한 대응)
- [ ] Auto 너비 개선 (한글 문자 너비 고려)
- [ ] 대용량 데이터 스트리밍 테스트 (10만 행)
- [ ] KDoc 주석 추가

---

## 모듈 구조

```
kotlin-excel-dsl/
├── excel-dsl/   # 통합 모듈 (권장)
├── core/        # 핵심 모델 & DSL (순수 Kotlin)
├── annotation/  # @Excel, @Column 처리
├── render/      # Apache POI 연동
└── theme/       # 미리 정의된 테마
```
