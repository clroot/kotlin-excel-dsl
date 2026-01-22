# Performance

[한국어](performance.ko.md)

The library uses Apache POI's SXSSF (Streaming Usermodel API) for memory-efficient large dataset handling.

## Benchmarks

**Test Environment**: MacBook Pro 14" M3 Pro (JMH benchmark with Sequence-based lazy data generation)

| Rows | Time | Peak Memory (min) | Peak Memory (avg) |
|------|------|-------------------|-------------------|
| 100,000 | ~0.4s | ~34 MB | ~101 MB |
| 500,000 | ~1.8s | ~51 MB | ~165 MB |
| 1,000,000 | ~3.5s | ~59 MB | ~134 MB |

## Key Features

### True Streaming

Data is processed row-by-row without loading the entire dataset into memory. This is achieved through Apache POI's SXSSF API, which only keeps a configurable window of rows in memory.

```kotlin
// Even with millions of rows, memory usage remains constant
excel {
    sheet<LargeData>("Data") {
        column("ID") { it.id }
        column("Value") { it.value }
        rows(largeDataSequence)  // Sequence is processed lazily
    }
}.writeTo(output)
```

### O(1) Memory for Auto-width

Column width calculation tracks only the maximum width encountered, not all cell values. This ensures auto-width doesn't become a memory bottleneck for large datasets.

### Near-constant Memory

Due to SXSSF streaming, increasing rows by 10x only increases peak memory by ~1.7x (min). The memory growth is primarily from POI's internal buffering, not from data storage.

## Configuration

The default row access window size is 100 rows. This can be adjusted if needed through the `PoiRenderer` constructor, though the default is suitable for most use cases.

## Best Practices

1. **Use Sequences for large datasets**: Pass `Sequence<T>` instead of `List<T>` to `rows()` for lazy evaluation.

2. **Avoid loading all data into memory**: Stream data from database or file sources directly.

3. **Consider fixed column widths**: For very large datasets, using fixed widths (`width = 20.chars`) is slightly more efficient than auto-width.
