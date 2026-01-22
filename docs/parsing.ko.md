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
