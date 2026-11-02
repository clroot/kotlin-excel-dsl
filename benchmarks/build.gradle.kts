plugins {
    kotlin("jvm")
    id("me.champeau.jmh") version "0.7.2"
}

dependencies {
    implementation(project(":core"))
    implementation(project(":render"))
    implementation(project(":theme"))

    jmh("org.openjdk.jmh:jmh-core:1.37")
    jmh("org.openjdk.jmh:jmh-generator-annprocess:1.37")
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
