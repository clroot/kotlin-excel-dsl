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

라이브러리는 다음 `ExcelException` 하위 타입을 던집니다.

| 예외 | 발생 시점 |
|------|-----------|
| `ExcelException` | 기본 타입. 라이브러리 오류 전체를 잡을 때 사용 |
| `ExcelConfigurationException` | 어노테이션·설정 오류: `@Excel`이나 `@Column` 누락, 제네릭 타입이 프로퍼티 타입과 맞지 않는 `ConditionalStyler`. `excelOf`/`parseExcel`에서 던지며, 스타일러 타입 불일치는 `writeTo` 중에 던집니다 |
| `ExcelWriteException` | 렌더링·쓰기 실패. 예: 값 추출 람다가 예외를 던지거나 출력 스트림이 실패한 경우. 데이터 셀에서 실패하면 `sheetName`, `rowIndex`, `columnIndex`가 채워지고 원인 예외가 보존됩니다 |
| `ExcelParseException` | 실패한 파싱 결과에서 `ParseResult.getOrThrow()`를 호출한 경우 |

렌더링 중 발생한 `ExcelConfigurationException`은 그대로 전달됩니다. 라이브러리 밖의 예외만 `ExcelWriteException`으로 감쌉니다.

`ExcelDataException`, `ColumnNotFoundException`, `StyleException`은 라이브러리가 던지지 않으므로 deprecated입니다. 파싱 중 컬럼 누락과 변환 실패는 `ParseResult.Failure`로 보고합니다([파싱](parsing.ko.md) 참고).

## 예외 계층 구조

```
ExcelException (기본)
├── ExcelConfigurationException
├── ExcelWriteException
└── ExcelParseException
```

라이브러리 예외는 모두 `ExcelException`을 상속하므로 catch 블록 하나로 처리할 수 있습니다:

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
    // 설정 오류 처리 (어노테이션 누락, 스타일러 타입 불일치 등)
} catch (e: ExcelWriteException) {
    // 렌더링·쓰기 오류 처리
    logger.error("실패 위치: sheet=${e.sheetName}, row=${e.rowIndex}, column=${e.columnIndex}", e)
}
```

`rowIndex`와 `columnIndex`는 0부터 시작하는 시트 기준 인덱스입니다(헤더 행 포함). 메시지에는 엑셀처럼 1부터 시작하는 번호로 표시됩니다:

```
ExcelWriteException: Failed to render cell: boom [sheet='Users', row=4, column=2]
```

## 인자 검증 (`IllegalArgumentException`)

DSL이나 어노테이션에 잘못된 인자를 넘기면 즉시 `IllegalArgumentException`을 던집니다. 이 예외는 `ExcelException`이 **아니므로** `catch (e: ExcelException)`으로 잡히지 않습니다:

- `freezePane(row, col)`에 음수 행이나 열을 넘긴 경우
- `formula("")`처럼 `Formula` 수식이 비어 있는 경우
- `backgroundColorHex`/`fontColorHex`의 `#RRGGBB` 값이 잘못된 경우 (`excelOf`에서 발생)
- `parseExcel`에 존재하지 않는 `sheetName`을 넘긴 경우

이 경우는 프로그래밍 오류이므로 잡기보다 호출하는 코드를 고칩니다.
