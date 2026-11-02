package io.clroot.excel.examples

import io.clroot.excel.annotation.Column
import io.clroot.excel.annotation.Excel
import io.clroot.excel.annotation.excelOf
import io.clroot.excel.core.dsl.excel
import io.clroot.excel.parser.ParseResult
import io.clroot.excel.parser.parseExcel
import io.clroot.excel.render.writeTo
import io.kotest.core.spec.style.StringSpec
import io.kotest.matchers.shouldBe
import java.io.ByteArrayOutputStream

@Excel
data class User(
    @Column("Name", order = 1) val name: String,
    @Column("Email", order = 2) val email: String,
)

class DocumentationTest : StringSpec({
    val users = listOf(User("Alice", "alice@example.com"), User("Bob", "bob@example.com"))

    "DSL and Sequence example round-trips through the parser" {
        val largeDataSequence = users.asSequence()
        val document =
            excel {
                sheet<User>("Users") {
                    column("Name") { it.name }
                    column("Email") { it.email }
                    rows(largeDataSequence.asIterable())
                }
            }
        val output = ByteArrayOutputStream()
        document.writeTo(output)
        output.toByteArray().inputStream().use { input ->
            parseExcel<User>(input) shouldBe ParseResult.Success(users)
        }
    }

    "annotation example round-trips through the parser" {
        val output = ByteArrayOutputStream()
        excelOf(users).writeTo(output)
        output.toByteArray().inputStream().use { input ->
            parseExcel<User>(input) shouldBe ParseResult.Success(users)
        }
    }
})
