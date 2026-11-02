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
    testImplementation("io.kotest:kotest-runner-junit5:6.0.7")
    testImplementation("io.kotest:kotest-assertions-core:6.0.7")
}
tasks.test { useJUnitPlatform() }
