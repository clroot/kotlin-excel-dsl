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
