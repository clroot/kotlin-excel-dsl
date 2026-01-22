# 사용 가이드

[English](guide.md)

Excel DSL의 스타일링, 고급 기능 및 상세 사용법을 다룹니다.

## 목차

- [스타일링](#스타일링)
  - [전역 스타일](#전역-스타일)
  - [컬럼별 스타일](#컬럼별-스타일)
  - [인라인 컬럼 스타일](#인라인-컬럼-스타일)
  - [헤더/바디 스타일 분리](#헤더바디-스타일-분리)
  - [조건부 스타일](#조건부-스타일)
  - [스타일 우선순위](#스타일-우선순위)
  - [테마](#테마)
- [어노테이션 스타일링](#어노테이션-스타일링)
  - [HeaderStyle과 BodyStyle](#headerstyle과-bodystyle)
  - [ConditionalStyle](#conditionalstyle)
  - [스타일 열거형](#스타일-열거형)
  - [HEX 색상 지원](#hex-색상-지원)
  - [어노테이션 스타일 우선순위](#어노테이션-스타일-우선순위)
- [고급 기능](#고급-기능)
  - [헤더 그룹](#헤더-그룹)
  - [컬럼 너비](#컬럼-너비)
  - [틀 고정](#틀-고정)
  - [자동 필터](#자동-필터)
  - [줄무늬 행 스타일](#줄무늬-행-스타일)
  - [수식](#수식)
  - [멀티 시트](#멀티-시트)

## 스타일링

### 전역 스타일

문서 전체의 헤더와 바디 셀에 스타일을 적용합니다:

```kotlin
excel {
    styles {
        header {
            backgroundColor(Color.GRAY)
            fontColor(Color.WHITE)
            bold()
            align(Alignment.CENTER)
        }
        body {
            border(BorderStyle.THIN)
        }
    }
    sheet<User>("사용자") {
        column("이름") { it.name }
        column("나이") { it.age }
        rows(users)
    }
}.writeTo(output)
```

### 컬럼별 스타일

`styles` DSL을 사용하여 특정 컬럼을 대상으로 스타일을 적용합니다:

```kotlin
excel {
    styles {
        header {
            backgroundColor(Color.GRAY)
            bold()
        }
        column("가격") {
            header {
                backgroundColor(Color.BLUE)
            }
            body {
                align(Alignment.RIGHT)
                numberFormat("#,##0")
            }
        }
    }
    sheet<Product>("상품") {
        column("상품명") { it.name }
        column("가격") { it.price }
        rows(products)
    }
}.writeTo(output)
```

### 인라인 컬럼 스타일

컬럼 정의 시 직접 스타일을 적용합니다:

```kotlin
excel {
    sheet<Product>("상품") {
        column("상품명") { it.name }
        column("가격", style = { align(Alignment.RIGHT); bold() }) { it.price }
        rows(products)
    }
}.writeTo(output)
```

### 헤더/바디 스타일 분리

컬럼의 헤더와 바디 셀에 각각 다른 스타일을 정의합니다:

```kotlin
excel {
    sheet<Product>("상품") {
        column("상품명") { it.name }
        column(
            "가격",
            headerStyle = { backgroundColor(Color.BLUE); bold() },
            bodyStyle = { align(Alignment.RIGHT) }
        ) { it.price }
        rows(products)
    }
}.writeTo(output)
```

### 조건부 스타일

렌더링 시점에 셀 값에 따라 동적으로 스타일을 적용합니다:

```kotlin
excel {
    sheet<Transaction>("거래내역") {
        column("내역") { it.description }
        column("금액", conditionalStyle = { value: Int? ->
            when {
                value == null -> null
                value < 0 -> CellStyleBuilder.fontColor(Color.RED)
                value > 1000000 -> CellStyleBuilder.fontColor(Color.GREEN)
                else -> null
            }
        }) { it.amount }
        rows(transactions)
    }
}.writeTo(output)
```

조건부 스타일은 바디 스타일과 함께 사용할 수 있습니다:

```kotlin
excel {
    sheet<Transaction>("거래내역") {
        column(
            "금액",
            bodyStyle = { bold(); align(Alignment.RIGHT) },
            conditionalStyle = { value: Int? ->
                if (value != null && value < 0) CellStyleBuilder.fontColor(Color.RED) else null
            }
        ) { it.amount }
        // 음수: bold + 우측 정렬 + 빨간색
        // 양수: bold + 우측 정렬
        rows(transactions)
    }
}.writeTo(output)
```

**편의 함수** - 간단한 경우에 사용합니다:

```kotlin
import io.clroot.excel.core.dsl.CellStyleBuilder.Companion.fontColor
import io.clroot.excel.core.dsl.CellStyleBuilder.Companion.backgroundColor

column("상태", conditionalStyle = { value: String? ->
    when (value) {
        "에러" -> fontColor(Color.RED)
        "성공" -> fontColor(Color.GREEN)
        else -> null
    }
}) { it.status }
```

### 스타일 우선순위

여러 스타일이 정의된 경우 다음 순서로 적용됩니다 (높은 우선순위 순):

1. **조건부 스타일** - 셀 값에 따른 동적 스타일
2. **인라인 스타일** - 컬럼에 직접 정의
3. **컬럼별 스타일** - `styles { column("이름") { ... } }`로 정의
4. **전역 스타일** - `styles { header { ... } }`로 정의
5. **줄무늬 행 스타일** - 짝수 인덱스 행에 적용

### 테마

미리 정의된 테마를 사용하여 일관된 스타일을 적용합니다:

```kotlin
excel(theme = Theme.Modern) {
    sheet<User>("사용자") {
        column("이름") { it.name }
        column("나이") { it.age }
        rows(users)
    }
}.writeTo(output)
```

| 테마 | 설명 |
|------|------|
| `Theme.Modern` | 파란 헤더, 흰색 글자, 가운데 정렬 |
| `Theme.Minimal` | 볼드 헤더, 배경색 없음 |
| `Theme.Classic` | 회색 헤더, 중간 두께 테두리 |

## 어노테이션 스타일링

어노테이션 방식(`@Excel`, `@Column`)을 사용할 때, `@HeaderStyle`, `@BodyStyle`, `@ConditionalStyle` 어노테이션으로 스타일을 적용할 수 있습니다.

### HeaderStyle과 BodyStyle

클래스 레벨(전체 컬럼) 또는 프로퍼티 레벨(특정 컬럼)에 스타일을 적용합니다:

```kotlin
@Excel
@HeaderStyle(bold = true, backgroundColor = StyleColor.LIGHT_GRAY)
@BodyStyle(alignment = StyleAlignment.CENTER)
data class User(
    @Column("이름", order = 1)
    @HeaderStyle(fontColor = StyleColor.BLUE)  // 클래스 레벨 오버라이드
    val name: String,

    @Column("나이", order = 2)
    @BodyStyle(bold = true)  // 클래스 레벨 오버라이드
    val age: Int,

    @Column("이메일", order = 3)
    val email: String  // 클래스 레벨 스타일 사용
)

excelOf(users).writeTo(output)
```

사용 가능한 속성:
- `bold` - 볼드체 (기본값: false)
- `italic` - 이탤릭체 (기본값: false)
- `backgroundColor` - `StyleColor` 열거형으로 배경색 지정
- `backgroundColorHex` - HEX 문자열로 배경색 지정 (예: "#FF5733")
- `fontColor` - `StyleColor` 열거형으로 글자색 지정
- `fontColorHex` - HEX 문자열로 글자색 지정
- `alignment` - `StyleAlignment` 열거형으로 정렬 지정
- `border` - `StyleBorder` 열거형으로 테두리 지정

### ConditionalStyle

`@ConditionalStyle` 어노테이션과 `ConditionalStyler<T>` 인터페이스를 사용하여 셀 값에 따른 동적 스타일을 적용합니다:

```kotlin
class ScoreStyler : ConditionalStyler<Int> {
    override fun style(value: Int?): CellStyle? = when {
        value == null -> null
        value >= 90 -> CellStyle(fontColor = Color.GREEN, bold = true)
        value >= 70 -> CellStyle(fontColor = Color.BLUE)
        value < 60 -> CellStyle(fontColor = Color.RED)
        else -> null
    }
}

class StatusStyler : ConditionalStyler<String> {
    override fun style(value: String?): CellStyle? = when (value) {
        "활성" -> CellStyle(backgroundColor = Color(200, 255, 200))
        "비활성" -> CellStyle(backgroundColor = Color(255, 200, 200))
        else -> null
    }
}

@Excel
data class Student(
    @Column("이름", order = 1)
    val name: String,

    @Column("점수", order = 2)
    @ConditionalStyle(ScoreStyler::class)
    val score: Int,

    @Column("상태", order = 3)
    @ConditionalStyle(StatusStyler::class)
    val status: String
)
```

**중요:** `ConditionalStyler<T>`의 제네릭 타입은 프로퍼티 타입과 일치해야 합니다. 타입이 맞지 않으면 런타임에 `ExcelConfigurationException`이 발생합니다.

### 스타일 열거형

타입 안전한 스타일 정의를 위한 열거형:

**StyleColor:**
```kotlin
enum class StyleColor {
    NONE,        // 색상 없음 (상속)
    WHITE, BLACK, GRAY, LIGHT_GRAY,
    RED, GREEN, BLUE, YELLOW, ORANGE
}
```

**StyleAlignment:**
```kotlin
enum class StyleAlignment {
    NONE,    // 정렬 없음 (상속)
    LEFT, CENTER, RIGHT
}
```

**StyleBorder:**
```kotlin
enum class StyleBorder {
    NONE,    // 테두리 없음 (상속)
    THIN, MEDIUM, THICK
}
```

### HEX 색상 지원

`StyleColor`에 없는 색상은 HEX 문자열로 지정합니다:

```kotlin
@Excel
@HeaderStyle(backgroundColorHex = "#4A90D9", fontColorHex = "#FFFFFF")
data class CustomStyled(
    @Column("값", order = 1)
    @BodyStyle(backgroundColorHex = "#F5F5F5")
    val value: String
)
```

지원 형식:
- 6자리 HEX: `"#RRGGBB"` (예: "#FF5733")
- # 없는 6자리: `"RRGGBB"` (예: "FF5733")

**참고:** 열거형 색상과 HEX 색상이 모두 지정된 경우, HEX가 우선 적용됩니다.

### 어노테이션 스타일 우선순위

스타일은 다음 순서로 적용됩니다 (높은 우선순위 순):

1. **ConditionalStyle** - `ConditionalStyler`의 동적 스타일
2. **프로퍼티 레벨 어노테이션** - 프로퍼티의 `@HeaderStyle`/`@BodyStyle`
3. **클래스 레벨 어노테이션** - 클래스의 `@HeaderStyle`/`@BodyStyle`
4. **테마 스타일** - `excelOf(data, theme = Theme.Modern)`에서 지정

예시:
```kotlin
@Excel
@HeaderStyle(bold = true)  // 우선순위 3
@BodyStyle(alignment = StyleAlignment.LEFT)  // 우선순위 3
data class Report(
    @Column("점수", order = 1)
    @BodyStyle(alignment = StyleAlignment.RIGHT)  // 우선순위 2 - 클래스 오버라이드
    @ConditionalStyle(ScoreStyler::class)  // 우선순위 1 - 조건부 추가
    val score: Int
)

// 테마와 함께 사용:
excelOf(reports, theme = Theme.Modern)  // 테마는 우선순위 4
```

## 고급 기능

### 헤더 그룹

셀 자동 병합을 지원하는 다중 헤더를 생성합니다:

```kotlin
excel {
    sheet<Student>("성적표") {
        headerGroup("학생 정보") {
            column("이름") { it.name }
            column("학번") { it.studentId }
        }
        headerGroup("성적") {
            column("국어") { it.korean }
            column("영어") { it.english }
        }
        rows(students)
    }
}.writeTo(output)
```

결과:

```
+-------학생 정보-------+--------성적--------+
+----이름----+---학번---+---국어---+---영어---+
|   김철수   |  20241   |    95    |    88    |
```

### 컬럼 너비

`chars` 확장 함수를 사용하여 고정 컬럼 너비를 설정합니다:

```kotlin
sheet<User>("사용자") {
    column("이름", width = 20.chars) { it.name }
    column("설명", width = 50.chars) { it.description }
    rows(users)
}
```

기본적으로 컬럼은 내용에 따라 최적의 너비를 계산하는 자동 너비를 사용합니다.

### 틀 고정

스크롤 시에도 특정 행/열이 항상 보이도록 고정합니다:

```kotlin
excel {
    sheet<User>("사용자") {
        freezePane(row = 1)              // 헤더 행 고정
        freezePane(row = 1, col = 1)     // 헤더 행 + 첫 번째 열 고정
        column("이름") { it.name }
        column("나이") { it.age }
        rows(users)
    }
}.writeTo(output)
```

파라미터:
- `row` - 상단에서 고정할 행 수 (기본값: 0)
- `col` - 왼쪽에서 고정할 열 수 (기본값: 0)

### 자동 필터

헤더 행에 필터 드롭다운을 추가합니다:

```kotlin
excel {
    sheet<User>("사용자") {
        autoFilter()
        column("이름") { it.name }
        column("나이") { it.age }
        column("부서") { it.department }
        rows(users)
    }
}.writeTo(output)
```

헤더 그룹이 있는 경우, 필터는 컬럼 헤더 행(두 번째 행)에 적용됩니다.

### 줄무늬 행 스타일

가독성 향상을 위한 교대 행 스타일을 적용합니다:

```kotlin
excel {
    sheet<User>("사용자") {
        alternateRowStyle {
            backgroundColor(Color.LIGHT_GRAY)
        }
        column("이름") { it.name }
        column("나이") { it.age }
        rows(users)
    }
}.writeTo(output)
```

스타일은 짝수 인덱스 데이터 행(0, 2, 4, ...)에 적용되며, 기존 body 스타일과 **병합**됩니다. 따라서 `alternateRowStyle`과 `bodyStyle`을 함께 사용할 수 있습니다:

```kotlin
excel {
    styles {
        body {
            bold()
            align(Alignment.CENTER)
        }
    }
    sheet<User>("사용자") {
        alternateRowStyle {
            backgroundColor(Color.LIGHT_GRAY)
        }
        // 짝수 행: bold + center + 회색 배경
        // 홀수 행: bold + center
        column("이름") { it.name }
        rows(users)
    }
}.writeTo(output)
```

### 수식

`formula()` 함수를 사용하여 셀에 Excel 수식을 지정합니다:

```kotlin
import io.clroot.excel.core.model.formula

excel {
    sheet<SummaryRow>("요약") {
        column("항목") { it.label }
        column("값") { formula("SUM(B2:B100)") }
        rows(summaryData)
    }
}.writeTo(output)
```

선행 `=` 기호는 선택사항입니다 - `formula("SUM(A1:A10)")`과 `formula("=SUM(A1:A10)")` 모두 동작합니다:

```kotlin
// 동일하게 동작
column("합계") { formula("SUM(A1:A10)") }
column("합계") { formula("=SUM(A1:A10)") }
```

다른 시트를 참조할 수도 있습니다:

```kotlin
column("총합") { formula("SUM(매출!B2:B100)") }
```

**참고:** 수식 유효성 검사는 Excel에서 파일을 열 때 수행됩니다. 잘못된 수식은 Excel에서 오류로 표시됩니다.

### 멀티 시트

여러 시트가 있는 워크북을 생성합니다:

```kotlin
excel {
    sheet<Product>("상품") {
        column("상품명") { it.name }
        column("가격") { it.price }
        rows(products)
    }
    sheet<Order>("주문") {
        column("주문번호") { it.id }
        column("수량") { it.amount }
        rows(orders)
    }
}.writeTo(output)
```
