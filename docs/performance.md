# Performance

[한국어](performance.ko.md)

Streaming based on Apache POI SXSSF for memory-efficient large dataset processing.

## Benchmarks

**Environment**
- MacBook Pro 14" M3 Pro
- JDK 21.0.8 (Amazon Corretto)
- JMH 1.37 (2 warmup, 3 measurement iterations, 10s each)

| Rows | Time | Memory (min / avg / max) |
|------|------|-------------------------|
| 100K | ~430ms | 34 / 106 / 189 MB |
| 500K | ~1.8s | 42 / 133 / 193 MB |
| 1M | ~3.5s | 59 / 136 / 193 MB |

Max memory stays constant at ~193MB even with 10x data increase.

## Key Features

- **Streaming rendering**: Row-by-row processing, only configured window (default 100 rows) kept in memory
- **O(1) Auto-width**: Tracks max width only, no cell value storage
- **Style caching**: Reuses identical styles

## Optimization Tips

```kotlin
// Pass large data as Sequence
rows(largeDataSequence)

// Fixed width is slightly more efficient
column("Amount", width = 15.chars) { it.amount }
```
