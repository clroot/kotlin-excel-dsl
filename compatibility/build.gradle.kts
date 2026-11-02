plugins {
    kotlin("jvm")
}
repositories {
    exclusiveContent {
        forRepository {
            maven { url = uri(providers.gradleProperty("artifactRepository").get()) }
        }
        filter { includeGroup("io.clroot.excel") }
    }
    mavenCentral()
}
kotlin {
    jvmToolchain(providers.gradleProperty("consumerJava").get().toInt())
}
dependencies {
    testImplementation("io.clroot.excel:excel-dsl:${providers.gradleProperty("artifactVersion").get()}")
    testImplementation(libs.kotest.runner)
    testImplementation(libs.kotest.assertions)
}
tasks.test { useJUnitPlatform() }
