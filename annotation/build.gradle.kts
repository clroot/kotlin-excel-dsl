plugins {
    kotlin("jvm")
}

description = "Annotation processing for kotlin-excel-dsl"

dependencies {
    api(project(":core"))
    implementation(kotlin("reflect"))
}
