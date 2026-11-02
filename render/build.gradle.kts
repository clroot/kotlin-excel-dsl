plugins {
    kotlin("jvm")
}

description = "Apache POI renderer for kotlin-excel-dsl"

dependencies {
    api(project(":core"))

    // Apache POI for Excel file generation
    implementation(libs.poi)
    implementation(libs.poi.ooxml)

    // Test dependencies
    testImplementation(project(":annotation"))
    testImplementation(project(":theme"))
}
