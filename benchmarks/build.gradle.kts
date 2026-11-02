plugins {
    kotlin("jvm")
    alias(libs.plugins.jmh)
}

dependencies {
    implementation(project(":core"))
    implementation(project(":render"))
    implementation(project(":theme"))

    jmh(libs.jmh.core)
    jmh(libs.jmh.generator)
}

jmh {
    warmupIterations = 2
    iterations = 3
    fork = 1
    warmup = "1s"
    timeOnIteration = "1s"
    resultFormat = "JSON"
    if (providers.gradleProperty("allocationProfile").isPresent) {
        profilers.add("gc")
    }
}
