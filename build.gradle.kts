import org.jetbrains.kotlin.gradle.dsl.JvmTarget
import org.jetbrains.kotlin.gradle.dsl.KotlinJvmProjectExtension

plugins {
    kotlin("jvm") version "2.2.0" apply false
    id("org.jlleitschuh.gradle.ktlint") version "12.2.0"
    `maven-publish`
    signing
}

allprojects {
    apply(plugin = "maven-publish")

    group = "io.clroot.excel"
    version = findProperty("version")?.toString()?.takeIf { it != "unspecified" } ?: "0.1.0-SNAPSHOT"

    repositories {
        mavenCentral()
    }
}

subprojects {
    apply(plugin = "org.jetbrains.kotlin.jvm")
    apply(plugin = "org.jlleitschuh.gradle.ktlint")
    apply(plugin = "maven-publish")
    apply(plugin = "signing")

    configure<JavaPluginExtension> {
        toolchain {
            languageVersion.set(JavaLanguageVersion.of(21))
        }
        withSourcesJar()
        withJavadocJar()
    }

    dependencies {
        "testImplementation"("io.kotest:kotest-runner-junit5:6.0.7")
        "testImplementation"("io.kotest:kotest-assertions-core:6.0.7")
    }

    tasks.withType<Test> {
        useJUnitPlatform()
    }

    configure<KotlinJvmProjectExtension> {
        compilerOptions {
            jvmTarget.set(JvmTarget.JVM_21)
            freeCompilerArgs.add("-Xjsr305=strict")
        }
    }

    configure<PublishingExtension> {
        publications {
            create<MavenPublication>("maven") {
                from(components["java"])

                pom {
                    name.set(project.name)
                    description.set(project.description ?: "Kotlin DSL for Excel file generation")
                    url.set("https://github.com/clroot/kotlin-excel-dsl")

                    licenses {
                        license {
                            name.set("MIT License")
                            url.set("https://opensource.org/licenses/MIT")
                        }
                    }

                    developers {
                        developer {
                            id.set("clroot")
                            name.set("clroot")
                            url.set("https://github.com/clroot")
                        }
                    }

                    scm {
                        connection.set("scm:git:git://github.com/clroot/kotlin-excel-dsl.git")
                        developerConnection.set("scm:git:ssh://github.com/clroot/kotlin-excel-dsl.git")
                        url.set("https://github.com/clroot/kotlin-excel-dsl")
                    }
                }
            }
        }
    }

    configure<SigningExtension> {
        val signingKey = findProperty("signingKey") as String? ?: System.getenv("GPG_SIGNING_KEY")
        val signingPassword = findProperty("signingPassword") as String? ?: System.getenv("GPG_SIGNING_PASSWORD")

        if (signingKey != null && signingPassword != null) {
            useInMemoryPgpKeys(signingKey, signingPassword)
        }

        sign(extensions.getByType<PublishingExtension>().publications["maven"])
    }

    tasks.withType<Sign>().configureEach {
        onlyIf {
            val signingKey = findProperty("signingKey") as String? ?: System.getenv("GPG_SIGNING_KEY")
            signingKey != null
        }
    }
}
