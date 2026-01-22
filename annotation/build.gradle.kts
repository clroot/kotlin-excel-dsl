plugins {
    kotlin("jvm")
}

description = "Annotation processing for kotlin-excel-dsl"

dependencies {
    api(project(":core"))
    implementation(kotlin("reflect"))

    testImplementation(project(":render"))
    testImplementation("org.apache.poi:poi-ooxml:5.5.1")
}
