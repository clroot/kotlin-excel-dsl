package io.clroot.excel.benchmarks

import io.clroot.excel.core.dsl.excel
import io.clroot.excel.render.writeTo
import org.openjdk.jmh.annotations.*
import org.openjdk.jmh.infra.Blackhole
import java.io.OutputStream
import java.time.LocalDate
import java.util.concurrent.TimeUnit

/**
 * Excel 생성 성능 벤치마크.
 *
 * 실행: ./gradlew :benchmarks:jmh
 */
@State(Scope.Benchmark)
@BenchmarkMode(Mode.AverageTime)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 2, time = 1)
@Measurement(iterations = 3, time = 1)
@Fork(1)
open class ExcelBenchmark {
    data class BenchmarkRow(
        val id: Long,
        val name: String,
        val value: Double,
        val date: LocalDate,
    )

    /**
     * 데이터를 Sequence로 lazy하게 생성.
     * 데이터 생성, DSL 구성, XLSX 렌더링 비용을 모두 측정한다.
     */
    private fun generateDataSequence(count: Int): Sequence<BenchmarkRow> =
        (1..count).asSequence().map { i ->
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
        ) {
        }
    }

    @Benchmark
    fun rows100k(blackhole: Blackhole) {
        val document =
            excel {
                sheet<BenchmarkRow>("Data") {
                    column("ID") { it.id }
                    column("Name") { it.name }
                    column("Value") { it.value }
                    column("Date") { it.date }
                    rows(generateDataSequence(100_000).asIterable())
                }
            }
        document.writeTo(NullOutputStream)
        blackhole.consume(document)
    }

    @Benchmark
    fun rows500k(blackhole: Blackhole) {
        val document =
            excel {
                sheet<BenchmarkRow>("Data") {
                    column("ID") { it.id }
                    column("Name") { it.name }
                    column("Value") { it.value }
                    column("Date") { it.date }
                    rows(generateDataSequence(500_000).asIterable())
                }
            }
        document.writeTo(NullOutputStream)
        blackhole.consume(document)
    }

    @Benchmark
    fun rows1m(blackhole: Blackhole) {
        val document =
            excel {
                sheet<BenchmarkRow>("Data") {
                    column("ID") { it.id }
                    column("Name") { it.name }
                    column("Value") { it.value }
                    column("Date") { it.date }
                    rows(generateDataSequence(1_000_000).asIterable())
                }
            }
        document.writeTo(NullOutputStream)
        blackhole.consume(document)
    }
}
