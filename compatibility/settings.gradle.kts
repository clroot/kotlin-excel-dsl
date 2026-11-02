pluginManagement {
    repositories {
        gradlePluginPortal()
        mavenCentral()
    }
    plugins {
        kotlin("jvm") version providers.gradleProperty("consumerKotlin").get()
    }
}
rootProject.name = "published-artifact-consumer"
