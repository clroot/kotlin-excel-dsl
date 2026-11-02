# Performance

[한국어](performance.ko.md)

Apache POI SXSSF streams rows with a default window of 100 rows and reuses styles.

## Method

Run timing and allocation profiling separately. GC profiler allocation counts cumulative allocated bytes, not peak heap or retained memory. Previous GC heap deltas and sampled memory figures are excluded from this baseline.

```bash
./gradlew :benchmarks:jmh --rerun-tasks
./gradlew :benchmarks:jmh -PallocationProfile --rerun-tasks
```

JMH 1.37 uses one fork, one thread, 2 × 1s warmup and 3 × 1s measurement iterations. Each operation includes lazy input generation, DSL construction and XLSX rendering for four columns and 100K/500K/1M rows. Output goes to a null stream; disk persistence is excluded. There is no forced GC, sleep or sampling thread inside the measured operation.

This short local baseline is not a performance guarantee or pass threshold. CI's 30-minute timeout prevents stalled jobs. For stable comparisons use the same hardware, JDK and heap settings with longer warmup and multiple forks.

## Sequence example

`rows` accepts `Iterable`, so adapt the sequence with `asIterable()` without materializing a list.

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

POI 6.0.0 measured 359.80 / 1716.10 / 3393.91 ms/op for 100K / 500K / 1M rows.
See the [comparison and raw results](../benchmarks/results/README.md#poi-6-comparison-2026-11-02)
for the observed 2–9% increases and measurement limitations.
