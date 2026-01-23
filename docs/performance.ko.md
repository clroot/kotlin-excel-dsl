# 성능

[English](performance.md)

Apache POI SXSSF 기반 스트리밍으로 대용량 데이터를 메모리 효율적으로 처리합니다.

## 벤치마크

**환경**
- MacBook Pro 14" M3 Pro
- JDK 21.0.8 (Amazon Corretto)
- JMH 1.37 (warmup 2회, measurement 3회, 각 10초)

| 행 수 | 처리 시간 | 메모리 (min / avg / max) |
|-------|----------|-------------------------|
| 100K | ~430ms | 34 / 106 / 189 MB |
| 500K | ~1.8초 | 42 / 133 / 193 MB |
| 1M | ~3.5초 | 59 / 136 / 193 MB |

데이터가 10배 증가해도 최대 메모리는 ~193MB로 일정합니다.

## 핵심 특징

- **스트리밍 렌더링**: 행 단위 처리, 설정된 윈도우(기본 100행)만 메모리 유지
- **O(1) Auto-width**: 최대 너비만 추적, 셀 값 저장 없음
- **스타일 캐싱**: 동일 스타일 재사용

## 최적화 팁

```kotlin
// 대용량 데이터는 Sequence로 전달
rows(largeDataSequence)

// 고정 너비 사용 시 약간 더 효율적
column("금액", width = 15.chars) { it.amount }
```
