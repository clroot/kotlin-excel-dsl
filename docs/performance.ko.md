# 성능

[English](performance.md)

Apache POI SXSSF로 행을 스트리밍합니다. 기본 윈도우는 100행이며 스타일을 재사용합니다.

## 측정 방법

생성 시간과 할당량을 별도로 실행합니다. GC 프로파일러의 할당량은 누적 할당 바이트이며 피크 힙이나 잔존 메모리가 아닙니다. 기존 GC 전후 힙 차이와 샘플링 결과는 비교 기준에서 제외했습니다.

```bash
./gradlew :benchmarks:jmh --rerun-tasks
./gradlew :benchmarks:jmh -PallocationProfile --rerun-tasks
```

JMH 1.37, fork 1개, 스레드 1개, warmup 2회 × 1초, 측정 3회 × 1초입니다. 각 작업은 4개 컬럼과 10만/50만/100만 행의 지연 데이터 생성, DSL 구성, XLSX 렌더링을 포함합니다. 출력은 null 스트림으로 보내며 디스크 저장 비용은 제외합니다. 측정 구간에 강제 GC·sleep·샘플링 스레드가 없습니다.

짧은 로컬 기준 측정이므로 성능 보장이나 합격 기준으로 사용하지 않습니다. CI의 30분 timeout은 작업 정지 방지용입니다. 안정적인 비교에는 같은 장비·JDK·힙 설정에서 더 긴 warmup과 여러 fork를 사용하세요.

## Sequence 예제

`rows`는 `Iterable`을 받으므로 `asIterable()`로 변환합니다. 전체 목록을 만들지 않습니다.

```kotlin
rows(largeDataSequence.asIterable())
```

[Executable DSL / annotation / parsing examples](../excel-dsl/src/test/kotlin/io/clroot/excel/examples/DocumentationTest.kt) · JDK 21+

[Raw results and environment](../benchmarks/results/README.md)

2026-11-02 · Intel Core Ultra 7 270K Plus · WSL2 x86_64 · Temurin 21.0.13+9

| Rows | ms/op (mean) |
|---|---:|
| 100k | 330.46 |
| 1m | 3268.07 |
| 500k | 1683.16 |

POI 6.0.0에서 10만/50만/100만 행은 각각 359.80 / 1716.10 / 3393.91 ms/op였습니다.
기존 대비 약 2~9% 증가했으나 짧은 측정만으로 회귀를 확정할 수 없습니다.
[비교 조건과 원본 결과](../benchmarks/results/README.md#poi-6-comparison-2026-11-02)를 참고하세요.
