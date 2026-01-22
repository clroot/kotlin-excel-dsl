# 에러 처리

[English](error-handling.md)

라이브러리는 힌트가 포함된 상세한 컨텍스트 기반 에러 메시지를 제공합니다.

## 에러 메시지

```kotlin
// @Excel 어노테이션 누락
excelOf(listOf(NonAnnotatedClass()))
// ExcelConfigurationException: Missing @Excel annotation [class=NonAnnotatedClass]
// Hint: Add @Excel annotation to your data class: @Excel data class NonAnnotatedClass(...)

// @Column 어노테이션 누락
excelOf(listOf(NoColumnClass()))
// ExcelConfigurationException: No properties annotated with @Column [class=NoColumnClass]
// Hint: Add @Column annotation to properties. Available properties: name, age
```

## 예외 타입

| 예외 | 설명 |
|------|------|
| `ExcelException` | 모든 Excel 관련 에러의 기본 예외 |
| `ExcelConfigurationException` | 설정 오류 (어노테이션, 설정값) |
| `ExcelDataException` | 데이터 처리 오류 |
| `ExcelWriteException` | 파일 쓰기 오류 |
| `ExcelParseException` | 엑셀 파싱 오류 |
| `ColumnNotFoundException` | 컬럼 조회 실패 |
| `StyleException` | 스타일 적용 오류 |

## 예외 계층 구조

```
ExcelException (기본)
├── ExcelConfigurationException
├── ExcelDataException
├── ExcelWriteException
├── ExcelParseException
├── ColumnNotFoundException
└── StyleException
```

모든 예외가 `ExcelException`을 상속하므로, 필요한 경우 단일 catch 블록으로 모든 라이브러리 에러를 처리할 수 있습니다:

```kotlin
try {
    excel {
        // ...
    }.writeTo(output)
} catch (e: ExcelException) {
    logger.error("Excel 작업 실패: ${e.message}")
}
```

또는 특정 예외만 처리할 수 있습니다:

```kotlin
try {
    excelOf(data).writeTo(output)
} catch (e: ExcelConfigurationException) {
    // 설정 오류 처리 (어노테이션 누락 등)
} catch (e: ExcelWriteException) {
    // 파일 쓰기 오류 처리
}
```
