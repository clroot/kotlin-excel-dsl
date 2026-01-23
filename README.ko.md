# kotlin-excel-dsl

[English](README.md)

타입 안전하고 우아한 문법으로 Excel 파일을 생성하는 Kotlin DSL 라이브러리.

## 특징

- **하이브리드 API**: 간단한 경우 어노테이션, 복잡한 경우 DSL
- **타입 안전성**: 컴파일 타임에 설정 오류 검증
- **선언적 스타일링**: 테마와 함께 직관적인 스타일 정의
- **어노테이션 스타일링**: `@HeaderStyle`, `@BodyStyle`, `@ConditionalStyle` 어노테이션 기반 스타일링
- **조건부 스타일**: 셀 값에 따른 동적 스타일 적용
- **수식 지원**: `formula("SUM(A1:A10)")` Excel 수식
- **헤더 그룹**: 셀 자동 병합을 지원하는 다중 헤더
- **틀 고정 / 자동 필터 / 줄무늬 행**: 일반적인 Excel 기능 지원
- **스트리밍 지원**: 대용량 데이터를 위한 SXSSF (100만+ 행)
- **엑셀 파싱**: 엑셀 파일을 데이터 클래스로 타입 안전하게 파싱

## 요구사항

- Kotlin 2.2.0+
- JDK 21+

## 설치

```kotlin
dependencies {
    implementation("io.clroot.excel:excel-dsl:$version")
}
```

> 최신 버전은 [Releases](https://github.com/clroot/kotlin-excel-dsl/releases)에서 확인하세요.

## 빠른 시작

### DSL 방식

```kotlin
data class User(val name: String, val age: Int, val joinedAt: LocalDate)

val users = listOf(
    User("김철수", 30, LocalDate.of(2024, 1, 15)),
    User("이영희", 25, LocalDate.of(2024, 3, 20))
)

excel {
    sheet<User>("사용자") {
        column("이름") { it.name }
        column("나이") { it.age }
        column("가입일") { it.joinedAt }
        rows(users)
    }
}.writeTo(FileOutputStream("users.xlsx"))
```

### 어노테이션 방식

```kotlin
@Excel
data class User(
    @Column("이름", order = 1) val name: String,
    @Column("나이", order = 2) val age: Int,
    @Column("가입일", order = 3) val joinedAt: LocalDate
)

excelOf(users).writeTo(FileOutputStream("users.xlsx"))
```

### 어노테이션 스타일링

```kotlin
@Excel
@HeaderStyle(bold = true, backgroundColor = StyleColor.LIGHT_GRAY)
@BodyStyle(alignment = StyleAlignment.CENTER)
data class StyledUser(
    @Column("이름", order = 1)
    @HeaderStyle(fontColor = StyleColor.BLUE)  // 프로퍼티 레벨 오버라이드
    val name: String,

    @Column("점수", order = 2)
    @ConditionalStyle(ScoreStyler::class)  // 동적 스타일링
    val score: Int
)

class ScoreStyler : ConditionalStyler<Int> {
    override fun style(value: Int?): CellStyle? = when {
        value == null -> null
        value >= 90 -> CellStyle(fontColor = Color.GREEN)
        value < 60 -> CellStyle(fontColor = Color.RED)
        else -> null
    }
}
```

### 테마 적용

```kotlin
excel(theme = Theme.Modern) {
    sheet<User>("사용자") {
        column("이름") { it.name }
        column("나이") { it.age }
        rows(users)
    }
}.writeTo(output)
```

### 조건부 스타일

```kotlin
excel {
    sheet<Transaction>("거래내역") {
        column("내역") { it.description }
        column("금액", conditionalStyle = { value: Int? ->
            when {
                value == null -> null
                value < 0 -> fontColor(Color.RED)
                value > 1000000 -> fontColor(Color.GREEN)
                else -> null
            }
        }) { it.amount }
        rows(transactions)
    }
}.writeTo(output)
```

### 수식

```kotlin
excel {
    sheet<SummaryRow>("요약") {
        column("항목") { it.label }
        column("값") { formula("SUM(A1:A10)") }  // Excel 수식
        rows(summaryData)
    }
}.writeTo(output)
```

### 엑셀 파싱

```kotlin
@Excel
data class User(
    @Column("이름") val name: String,
    @Column("이메일") val email: String,
)

val result = parseExcel<User>(FileInputStream("users.xlsx"))
when (result) {
    is ParseResult.Success -> result.data.forEach { println(it) }
    is ParseResult.Failure -> result.errors.forEach { println(it.message) }
}
```

## 문서

- [사용 가이드](docs/guide.ko.md) - 스타일링, 테마, 고급 기능
- [파싱](docs/parsing.ko.md) - 엑셀 파일 파싱 설정
- [에러 처리](docs/error-handling.ko.md) - 예외 타입 및 처리
- [성능](docs/performance.ko.md) - 벤치마크 및 모범 사례

## 모듈 구조

```
kotlin-excel-dsl/
├── excel-dsl/   # 통합 모듈 (권장)
├── core/        # 핵심 모델 & DSL
├── annotation/  # @Excel, @Column 처리
├── render/      # Apache POI 연동
├── theme/       # 미리 정의된 테마
└── parser/      # 엑셀 파일 파싱
```

## 라이선스

MIT License
