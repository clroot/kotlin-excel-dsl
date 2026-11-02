import org.jetbrains.kotlin.gradle.dsl.JvmTarget
import org.jetbrains.kotlin.gradle.dsl.KotlinJvmProjectExtension
import org.jetbrains.kotlin.gradle.dsl.abi.ExperimentalAbiValidation

plugins {
    alias(libs.plugins.kotlin.jvm) apply false
    alias(libs.plugins.ktlint)
    `maven-publish`
    signing
}

val catalog = libs

val publishedModules = setOf("core", "annotation", "render", "parser", "theme", "excel-dsl")

allprojects {

    // Preserve the existing lint rules while upgrading the Gradle integration.
    pluginManager.withPlugin("org.jlleitschuh.gradle.ktlint") {
        configure<org.jlleitschuh.gradle.ktlint.KtlintExtension> {
            version.set(catalog.versions.ktlint.asProvider().get())
        }
    }

    group = "io.clroot.excel"
    version = findProperty("version")?.toString()?.takeIf { it != "unspecified" } ?: "0.1.0-SNAPSHOT"

    repositories {
        mavenCentral()
    }
}

subprojects {
    apply(plugin = "org.jetbrains.kotlin.jvm")
    apply(plugin = "org.jlleitschuh.gradle.ktlint")

    configure<JavaPluginExtension> {
        toolchain {
            languageVersion.set(JavaLanguageVersion.of(21))
        }
        withSourcesJar()
        withJavadocJar()
    }

    dependencies {
        "testImplementation"(catalog.kotest.runner)
        "testImplementation"(catalog.kotest.assertions)
    }

    val testJavaVersion = providers.gradleProperty("testJavaVersion").map(String::toInt).orElse(21)
    val javaToolchains = extensions.getByType<JavaToolchainService>()
    tasks.withType<Test>().configureEach {
        useJUnitPlatform()
        javaLauncher.set(
            javaToolchains.launcherFor {
                languageVersion.set(testJavaVersion.map(JavaLanguageVersion::of))
            },
        )
        // Print the worker JVM's actual version, even without Gradle --info.
        jvmArgs("-showversion")
        doFirst {
            val launcher = javaLauncher.get()
            logger.lifecycle("Test JVM: ${launcher.executablePath} (${launcher.metadata.languageVersion})")
        }
    }

    configure<KotlinJvmProjectExtension> {
        compilerOptions {
            jvmTarget.set(JvmTarget.JVM_21)
            freeCompilerArgs.add("-Xjsr305=strict")
        }
    }

    if (name !in publishedModules) return@subprojects

    configure<KotlinJvmProjectExtension> {
        @OptIn(ExperimentalAbiValidation::class)
        extensions.configure<org.jetbrains.kotlin.gradle.dsl.abi.AbiValidationExtension> { enabled.set(true) }
    }

    apply(plugin = "maven-publish")
    apply(plugin = "signing")

    configure<PublishingExtension> {
        repositories {
            maven {
                name = "verification"
                url = rootProject.layout.buildDirectory.dir("verification-repository").get().asFile.toURI()
            }
        }
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
