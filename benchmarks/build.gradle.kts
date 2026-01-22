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
    // 메모리 측정을 위한 GC 프로파일러
    profilers.add("gc")
}
