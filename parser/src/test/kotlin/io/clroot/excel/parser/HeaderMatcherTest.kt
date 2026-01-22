package io.clroot.excel.parser

import io.kotest.core.spec.style.DescribeSpec
import io.kotest.matchers.shouldBe

class HeaderMatcherTest : DescribeSpec({

    describe("HeaderMatcher.EXACT") {
        val matcher = HeaderMatcher(HeaderMatching.EXACT)

        it("정확히 일치하면 true") {
            matcher.matches("이름", "이름", emptyArray()) shouldBe true
        }

        it("대소문자가 다르면 false") {
            matcher.matches("Name", "name", emptyArray()) shouldBe false
        }

        it("공백이 다르면 false") {
            matcher.matches("이 름", "이름", emptyArray()) shouldBe false
        }

        it("alias가 정확히 일치하면 true") {
            matcher.matches("Name", "이름", arrayOf("Name", "성명")) shouldBe true
        }
    }

    describe("HeaderMatcher.FLEXIBLE") {
        val matcher = HeaderMatcher(HeaderMatching.FLEXIBLE)

        it("정확히 일치하면 true") {
            matcher.matches("이름", "이름", emptyArray()) shouldBe true
        }

        it("대소문자가 달라도 true") {
            matcher.matches("Name", "name", emptyArray()) shouldBe true
            matcher.matches("NAME", "name", emptyArray()) shouldBe true
        }

        it("앞뒤 공백을 무시한다") {
            matcher.matches("  이름  ", "이름", emptyArray()) shouldBe true
        }

        it("연속 공백을 단일 공백으로 정규화한다") {
            matcher.matches("이  름", "이 름", emptyArray()) shouldBe true
        }

        it("alias도 유연하게 매칭한다") {
            matcher.matches("  name  ", "이름", arrayOf("Name", "성명")) shouldBe true
        }

        it("매칭되지 않으면 false") {
            matcher.matches("나이", "이름", emptyArray()) shouldBe false
        }
    }
})
