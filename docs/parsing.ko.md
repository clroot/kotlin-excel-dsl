# 엑셀 파싱

[English](parsing.md)

동일한 어노테이션을 사용하여 엑셀 파일을 타입 안전한 데이터 클래스로 파싱합니다.

## 목차

- [기본 파싱](#기본-파싱)
- [헤더 별칭](#헤더-별칭)
- [파싱 설정](#파싱-설정)
- [설정 옵션](#설정-옵션)
- [지원 타입](#지원-타입)

## 기본 파싱

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

## 헤더 별칭

유연한 매칭을 위한 대체 헤더명을 지원합니다:

```kotlin
@Excel
data class User(
    @Column("이름", aliases = ["Name", "성명"]) val name: String,
    @Column("이메일", aliases = ["Email", "E-mail"]) val email: String,
)
```

이를 통해 다양한 헤더 명명 규칙을 가진 엑셀 파일을 파싱할 수 있습니다.

## 파싱 설정

```kotlin
val result = parseExcel<User>(inputStream) {
    // 시트 선택
    sheetName = "사용자"       // 또는 sheetIndex = 0
    headerRow = 1              // 0부터 시작
    
    // 헤더 매칭
    headerMatching = HeaderMatching.FLEXIBLE  // 대소문자 무시, 공백 정규화
    
    // 에러 처리
    onError = OnError.COLLECT  // 모든 에러 수집, 또는 FAIL_FAST
    
    // 데이터 처리
    skipEmptyRows = true
    trimWhitespace = true
    treatBlankAsNull = true
    
    // 검증
    validateRow { user -> 
        require(user.email.contains("@")) { "이메일 형식 오류" }
    }
    validateAll { users ->
        val duplicates = users.groupBy { it.email }.filter { it.value.size > 1 }
        require(duplicates.isEmpty()) { "중복 이메일: ${duplicates.keys}" }
    }
    
    // 커스텀 타입 컨버터
    converter<Money> { value -> Money(BigDecimal(value.toString())) }
}
```

## 설정 옵션

| 옵션 | 타입 | 기본값 | 설명 |
|------|------|--------|------|
| `sheetIndex` | `Int` | `0` | 시트 인덱스 (0부터 시작) |
| `sheetName` | `String?` | `null` | 시트 이름 (설정 시 sheetIndex보다 우선) |
| `headerRow` | `Int` | `0` | 헤더 행 인덱스 |
| `headerMatching` | `HeaderMatching` | `FLEXIBLE` | `EXACT` 또는 `FLEXIBLE` |
| `onError` | `OnError` | `COLLECT` | `COLLECT` 또는 `FAIL_FAST` |
| `skipEmptyRows` | `Boolean` | `true` | 빈 행 건너뛰기 |
| `trimWhitespace` | `Boolean` | `true` | 문자열 앞뒤 공백 제거 |
| `treatBlankAsNull` | `Boolean` | `true` | 공백 문자열을 null로 처리 |

### 헤더 매칭 모드

- **EXACT**: 헤더가 정확히 일치해야 함 (대소문자 구분)
- **FLEXIBLE**: 대소문자 무시, 공백 정규화 (권장)

### 에러 처리 모드

- **COLLECT**: 모든 에러를 수집하여 `ParseResult.Failure`로 반환
- **FAIL_FAST**: 첫 번째 에러에서 중단

## 지원 타입

다음 타입들이 기본적으로 지원됩니다:

- `String`
- `Int`, `Long`
- `Double`, `Float`
- `Boolean`
- `BigDecimal`
- `LocalDate`
- `LocalDateTime`

커스텀 타입의 경우 `converter` 함수를 사용합니다:

```kotlin
data class Money(val amount: BigDecimal)

val result = parseExcel<Product>(inputStream) {
    converter<Money> { value -> Money(BigDecimal(value.toString())) }
}
```

## 정수와 Boolean 변환 정책

기본 Int/Long 변환기는 문자열을 Double을 거치지 않고 파싱합니다. 값은 수학적으로 정수이며 대상 타입의 범위 안에 있어야 합니다. `42.0`, `4.2e1` 같은 소수점·지수 표기는 허용하지만, 소수 부분이 있는 값, 범위 초과, NaN, Infinity는 IllegalArgumentException으로 처리합니다. 숫자 셀은 저장된 값을 기준으로 검사하므로 Excel에서 이미 손실된 정밀도는 복구할 수 없습니다. 큰 정수 식별자는 텍스트 셀로 저장하세요.

Boolean은 Boolean 값, 대소문자를 구분하지 않는 `true`/`false` 문자열, 숫자 0/1만 허용합니다. 문자열 `"0"`/`"1"`, 그 밖의 숫자와 잘못된 문자열은 오류입니다. 공백 제거와 nullable/빈 셀 처리는 기존 설정을 따릅니다. 사용자 지정 converter는 계속 우선 적용되며 원본 값을 받습니다.

기존 동작과 호환성이 달라집니다. 소수 부분을 버리거나 범위 초과를 보정하지 않으며, 잘못된 Boolean 문자열을 false로 처리하지 않습니다. 원본 데이터를 정리하거나 기존 정책이 필요한 경우 converter를 명시적으로 등록하세요.

```kotlin
parseExcel<MyRow>(input) {
    converter<Int> { value -> value.toString().toDouble().toInt() }
    converter<Boolean> { value ->
        when (value) {
            is Boolean -> value
            is Number -> value.toInt() != 0
            else -> value.toString().trim().equals("true", ignoreCase = true)
        }
    }
}
```

위 전환 예시는 값 손실을 허용하는 변환을 의도적으로 복원합니다. 변환 실패는 기존처럼 행·열 정보가 담긴 ParseError로 보고됩니다. FAIL_FAST는 첫 실패 행에서 반환하고, COLLECT는 여러 행의 오류를 모읍니다.

DSL·Sequence·어노테이션·파싱의 [실행 가능한 공통 예제](../excel-dsl/src/test/kotlin/io/clroot/excel/examples/DocumentationTest.kt)는 CI의 `:excel-dsl:test`에서 검증합니다. JDK 21 이상이 필요합니다.
