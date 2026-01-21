plugins {
    kotlin("jvm")
}

description = "Excel parser for kotlin-excel-dsl"

dependencies {
    api(project(":core"))
    api(project(":annotation"))

    // Apache POI for Excel file reading
    implementation("org.apache.poi:poi:5.5.1")
    implementation("org.apache.poi:poi-ooxml:5.5.1")
    implementation(kotlin("reflect"))
}
