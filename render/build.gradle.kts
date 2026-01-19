plugins {
    kotlin("jvm")
}

description = "Apache POI renderer for kotlin-excel-dsl"

dependencies {
    api(project(":core"))

    // Apache POI for Excel file generation
    implementation("org.apache.poi:poi:5.2.5")
    implementation("org.apache.poi:poi-ooxml:5.2.5")

    // Test dependencies
    testImplementation(project(":annotation"))
    testImplementation(project(":theme"))
}
