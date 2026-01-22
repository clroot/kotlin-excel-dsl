pluginManagement {
    repositories {
        mavenCentral()
        gradlePluginPortal()
    }
}

plugins {
    id("com.gradleup.nmcp.settings") version "1.4.3"
}

rootProject.name = "kotlin-excel-dsl"

include("core")
include("annotation")
include("render")
include("parser")
include("theme")
include("excel-dsl")
include("benchmarks")

nmcpSettings {
    centralPortal {
        username = providers.environmentVariable("SONATYPE_USERNAME").orNull
            ?: providers.gradleProperty("sonatypeUsername").orNull
        password = providers.environmentVariable("SONATYPE_PASSWORD").orNull
            ?: providers.gradleProperty("sonatypePassword").orNull
        publishingType = "AUTOMATIC"
    }
}
