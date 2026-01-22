# 사용 가이드

[English](guide.md)

Excel DSL의 스타일링, 고급 기능 및 상세 사용법을 다룹니다.

## 목차

- [스타일링](#스타일링)
  - [전역 스타일](#전역-스타일)
  - [컬럼별 스타일](#컬럼별-스타일)
  - [인라인 컬럼 스타일](#인라인-컬럼-스타일)
  - [헤더/바디 스타일 분리](#헤더바디-스타일-분리)
  - [스타일 우선순위](#스타일-우선순위)
  - [테마](#테마)
- [고급 기능](#고급-기능)
  - [헤더 그룹](#헤더-그룹)
  - [컬럼 너비](#컬럼-너비)
  - [틀 고정](#틀-고정)
  - [자동 필터](#자동-필터)
  - [줄무늬 행 스타일](#줄무늬-행-스타일)
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

### 스타일 우선순위

여러 스타일이 정의된 경우 다음 순서로 적용됩니다 (높은 우선순위 순):

1. **인라인 스타일** - 컬럼에 직접 정의
2. **컬럼별 스타일** - `styles { column("이름") { ... } }`로 정의
3. **전역 스타일** - `styles { header { ... } }`로 정의

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
