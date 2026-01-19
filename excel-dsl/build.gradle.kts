plugins {
    kotlin("jvm")
}

description = "All-in-one module for kotlin-excel-dsl"

dependencies {
    api(project(":core"))
    api(project(":annotation"))
    api(project(":render"))
    api(project(":theme"))
}
