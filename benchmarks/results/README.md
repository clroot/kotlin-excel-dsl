# Local benchmark baseline

Date: 2026-11-02. Source: issue #18 working tree (benchmark implementation committed alongside these results).

- CPU: Intel Core Ultra 7 270K Plus, 24 logical CPUs
- OS: Linux 6.18.54.1-microsoft-standard-WSL2, x86_64
- Benchmark JVM: Temurin 21.0.13+9-LTS, no explicit heap settings (complete JVM arguments in JSON)
- Gradle: 9.1.0; Kotlin: 2.3.21; JMH: 1.37
- One thread, one fork, two 1-second warmup iterations, three 1-second measurement iterations. A complete operation can exceed the iteration duration.
- Four columns (Long, String, Double, LocalDate); 100K, 500K, 1M lazily generated rows. Includes input construction, DSL and XLSX serialization to a null stream. SXSSF temporary file I/O remains included.
- Timing and allocation runs were sequential, after build verification. Other host activity is uncontrolled. These short runs are exploratory, not portable performance targets.

Commands:

```bash
./gradlew :benchmarks:jmh --rerun-tasks
cp benchmarks/build/results/jmh/results.json benchmarks/results/timing.json
./gradlew :benchmarks:jmh -PallocationProfile --rerun-tasks
cp benchmarks/build/results/jmh/results.json benchmarks/results/allocation.json
```

[timing.json](timing.json) and [allocation.json](allocation.json) retain per-iteration raw results and JVM metadata. The corresponding `*-output.txt` logs retain the full invocation output. Allocation profiling reports cumulative allocated bytes per operation and GC activity, not peak/live/retained heap. Profiler timing is not substituted for the unprofiled baseline. Previous sampled peak figures are not comparable.


| Rows | ms/op (mean) |
|---|---:|
| 100k | 330.46 |
| 1m | 3268.07 |
| 500k | 1683.16 |


## POI 6 comparison (2026-11-02)

Same machine, JDK 21, Kotlin 2.3.21 and JMH settings as the baseline; sequential
timing and allocation runs. POI changed from 5.5.1 to 6.0.0.

| Rows | POI 5.5.1 ms/op | POI 6.0.0 ms/op | Change |
|---|---:|---:|---:|
| 100k | 330.46 | 359.80 | +8.9% |
| 500k | 1683.16 | 1716.10 | +2.0% |
| 1m | 3268.07 | 3393.91 | +3.9% |

These short runs have uncontrolled host activity and broad uncertainty; the
observed increases do not establish a statistically reliable regression.
Use longer warmup and multiple forks before setting a performance threshold.
Raw results: [timing](poi6-timing.json), [allocation](poi6-allocation.json).
Full invocation output is retained in the corresponding poi6-*-output.txt files.

GC profiler cumulative allocation (bytes/op, not peak heap):

| Rows | POI 5.5.1 | POI 6.0.0 |
|---|---:|---:|
| rows100k | 287275252 | 291105618 |
| rows1m | 2864243072 | 2925406723 |
| rows500k | 1450226056 | 1450056576 |
