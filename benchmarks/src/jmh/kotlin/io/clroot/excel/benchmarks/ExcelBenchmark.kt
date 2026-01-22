package io.clroot.excel.benchmarks

import io.clroot.excel.core.dsl.excel
import io.clroot.excel.render.writeTo
import org.openjdk.jmh.annotations.*
import org.openjdk.jmh.infra.Blackhole
import java.io.OutputStream
import java.lang.management.ManagementFactory
import java.time.LocalDate
import java.util.concurrent.TimeUnit
import java.util.concurrent.atomic.AtomicBoolean
import java.util.concurrent.atomic.AtomicLong

/**
 * Excel 생성 성능 벤치마크.
 *
 * 실행: ./gradlew :benchmarks:jmh
 */
@State(Scope.Benchmark)
@BenchmarkMode(Mode.AverageTime)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 5, time = 10)
@Measurement(iterations = 10, time = 10)
@Fork(1)
open class ExcelBenchmark {
    data class BenchmarkRow(
        val id: Long,
        val name: String,
        val value: Double,
        val date: LocalDate,
    )

    private lateinit var data100k: List<BenchmarkRow>
    private lateinit var data500k: List<BenchmarkRow>
    private lateinit var data1m: List<BenchmarkRow>

    private val memoryMXBean = ManagementFactory.getMemoryMXBean()

    // 피크 메모리 결과 저장
    private val peakMemoryResults = mutableMapOf<String, MutableList<Double>>()

    @Setup
    fun setup() {
        data100k = generateData(100_000)
        data500k = generateData(500_000)
        data1m = generateData(1_000_000)

        // GC로 데이터 생성 메모리 정리
        repeat(3) { System.gc() }
        Thread.sleep(500)
    }

    @TearDown(Level.Trial)
    fun tearDown() {
        // 최종 결과 출력
        if (peakMemoryResults.isNotEmpty()) {
            println("\n========== Peak Memory Summary ==========")
            peakMemoryResults.toSortedMap().forEach { (label, values) ->
                if (values.isNotEmpty()) {
                    val avg = values.average()
                    val min = values.minOrNull() ?: 0.0
                    val max = values.maxOrNull() ?: 0.0
                    println(
                        "$label rows: avg=${String.format("%.1f", avg)}MB, min=${
                            String.format(
                                "%.1f",
                                min,
                            )
                        }MB, max=${String.format("%.1f", max)}MB (n=${values.size})",
                    )
                }
            }
            println("==========================================\n")
            peakMemoryResults.clear()
        }
    }

    private fun generateData(count: Int): List<BenchmarkRow> =
        (1..count).map { i ->
            BenchmarkRow(
                id = i.toLong(),
                name = "Name-$i",
                value = i * 1.5,
                date = LocalDate.of(2024, 1, 1).plusDays(i.toLong() % 365),
            )
        }

    private object NullOutputStream : OutputStream() {
        override fun write(b: Int) {}

        override fun write(b: ByteArray) {}

        override fun write(
            b: ByteArray,
            off: Int,
            len: Int,
        ) {}
    }

    private inline fun measurePeakMemory(
        label: String,
        block: () -> Unit,
    ) {
        // GC로 베이스라인 정리
        System.gc()
        Thread.sleep(50)

        val baselineMemory = memoryMXBean.heapMemoryUsage.used
        val peakMemory = AtomicLong(baselineMemory)
        val running = AtomicBoolean(true)

        // 백그라운드 메모리 샘플링 스레드
        val samplerThread =
            Thread {
                try {
                    while (running.get()) {
                        val used = memoryMXBean.heapMemoryUsage.used
                        peakMemory.updateAndGet { maxOf(it, used) }
                        Thread.sleep(5)
                    }
                } catch (_: InterruptedException) {
                    // 정상 종료
                }
            }
        samplerThread.isDaemon = true
        samplerThread.start()

        try {
            block()
        } finally {
            running.set(false)
            samplerThread.join(100)
        }

        val peakUsedMB = (peakMemory.get() - baselineMemory) / 1_000_000.0
        peakMemoryResults.getOrPut(label) { mutableListOf() }.add(peakUsedMB)
    }

    @Benchmark
    fun rows100k(blackhole: Blackhole) {
        measurePeakMemory("100k") {
            val document =
                excel {
                    sheet<BenchmarkRow>("Data") {
                        column("ID") { it.id }
                        column("Name") { it.name }
                        column("Value") { it.value }
                        column("Date") { it.date }
                        rows(data100k)
                    }
                }
            document.writeTo(NullOutputStream)
            blackhole.consume(document)
        }
    }

    @Benchmark
    fun rows500k(blackhole: Blackhole) {
        measurePeakMemory("500k") {
            val document =
                excel {
                    sheet<BenchmarkRow>("Data") {
                        column("ID") { it.id }
                        column("Name") { it.name }
                        column("Value") { it.value }
                        column("Date") { it.date }
                        rows(data500k)
                    }
                }
            document.writeTo(NullOutputStream)
            blackhole.consume(document)
        }
    }

    @Benchmark
    fun rows1m(blackhole: Blackhole) {
        measurePeakMemory("1m") {
            val document =
                excel {
                    sheet<BenchmarkRow>("Data") {
                        column("ID") { it.id }
                        column("Name") { it.name }
                        column("Value") { it.value }
                        column("Date") { it.date }
                        rows(data1m)
                    }
                }
            document.writeTo(NullOutputStream)
            blackhole.consume(document)
        }
    }
}
