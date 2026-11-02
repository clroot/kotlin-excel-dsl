plugins {
    kotlin("jvm")
}

description = "Excel parser for kotlin-excel-dsl"

dependencies {
    api(project(":core"))
    api(project(":annotation"))

    // Apache POI for Excel file reading
    implementation(libs.poi)
    implementation(libs.poi.ooxml)
    implementation(kotlin("reflect"))

    // Test dependencies for roundtrip tests
    testImplementation(project(":render"))
}
