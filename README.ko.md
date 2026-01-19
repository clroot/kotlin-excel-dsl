# kotlin-excel-dsl

[English](README.md)

타입 안전하고 우아한 문법으로 Excel 파일을 생성하는 Kotlin DSL 라이브러리.

## 요구사항

- Kotlin 2.2.0+
- JDK 21+

## 특징

- **하이브리드 API**: 간단한 경우 어노테이션, 복잡한 경우 DSL
- **타입 안전성**: 컴파일 타임에 설정 오류 검증
- **CSS-like 스타일링**: 컬럼 단위 세밀한 스타일 제어
- **미리 정의된 테마**: Modern, Minimal, Classic
- **헤더 그룹**: 셀 자동 병합을 지원하는 다중 헤더
- **날짜 자동 포맷**: LocalDate/LocalDateTime 자동 감지
- **스트리밍 지원**: 대용량 데이터를 위한 SXSSF
- **상세한 에러 메시지**: 힌트가 포함된 컨텍스트 기반 예외

## 빠른 시작

### Gradle

```kotlin
dependencies {
    // 통합 모듈 (권장)
    implementation("io.clroot.excel:excel-dsl:0.1.0")
}
```

개별 모듈 사용:

```kotlin
dependencies {
    implementation("io.clroot.excel:core:0.1.0")
    implementation("io.clroot.excel:annotation:0.1.0")
    implementation("io.clroot.excel:render:0.1.0")
    implementation("io.clroot.excel:theme:0.1.0")
}
```

## 사용법

```kotlin
import io.clroot.excel.*
```

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
    @Column("이름", order = 1)
    val name: String,
    
    @Column("나이", order = 2)
    val age: Int,
    
    @Column("가입일", order = 3)
    val joinedAt: LocalDate
)

excelOf(users).writeTo(FileOutputStream("users.xlsx"))
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

### 커스텀 스타일

#### 전역 스타일

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

#### 컬럼별 스타일 (styles DSL)

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

#### 인라인 컬럼 스타일

```kotlin
excel {
    sheet<Product>("상품") {
        column("상품명") { it.name }
        column("가격", style = { align(Alignment.RIGHT); bold() }) { it.price }
        rows(products)
    }
}.writeTo(output)
```

#### 컬럼별 헤더/바디 스타일 분리

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

**스타일 우선순위**: 인라인 > 컬럼별(styles DSL) > 전역

### 헤더 그룹 (다중 헤더)

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

```kotlin
sheet<User>("사용자") {
    column("이름", width = 20.chars) { it.name }
    column("설명", width = 50.chars) { it.description }
    rows(users)
}
```

### 멀티 시트

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

## 에러 처리

상세한 컨텍스트가 포함된 에러 메시지를 제공합니다:

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

### 예외 타입

| 예외                          | 설명                                     |
| ----------------------------- | ---------------------------------------- |
| `ExcelException`              | 모든 Excel 관련 에러의 기본 예외         |
| `ExcelConfigurationException` | 설정 오류 (어노테이션, 설정값)           |
| `ExcelDataException`          | 데이터 처리 오류                         |
| `ExcelWriteException`         | 파일 쓰기 오류                           |
| `ColumnNotFoundException`     | 컬럼 조회 실패                           |
| `StyleException`              | 스타일 적용 오류                         |

## 제공 테마

| 테마 | 설명 |
|------|------|
| `Theme.Modern` | 파란 헤더, 흰색 글자, 가운데 정렬 |
| `Theme.Minimal` | 볼드 헤더, 배경색 없음 |
| `Theme.Classic` | 회색 헤더, 중간 두께 테두리 |

## 모듈 구조

```
kotlin-excel-dsl/
├── excel-dsl/   # 통합 모듈 (권장)
├── core/        # 핵심 모델 & DSL (순수 Kotlin)
├── annotation/  # @Excel, @Column 처리
├── render/      # Apache POI 연동
└── theme/       # 미리 정의된 테마
```

## 라이선스

MIT License
