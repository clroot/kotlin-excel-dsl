# Development

Install JDK 21 (required for compilation). To verify Java 25 compatibility, also
install JDK 25. The Gradle Wrapper downloads Gradle 9.1.0 and verifies its SHA-256;
no system Gradle installation is needed.

| Purpose | Supported JDK |
| --- | --- |
| Gradle runtime | 21 or 25 |
| Compilation toolchain | 21 |
| Kotlin JVM target / Java bytecode | 21 |
| Test runtime | 21 by default; 25 with `-PtestJavaVersion=25` |
| Library minimum runtime | 21 |

Clone the repository and run from its root:

```bash
git clone https://github.com/clroot/kotlin-excel-dsl.git
cd kotlin-excel-dsl
export JAVA_HOME=/absolute/path/to/jdk-21
export PATH="$JAVA_HOME/bin:$PATH"
./gradlew --version
./gradlew build --warning-mode=fail
```

Replace the JDK path with your installed JDK directory. This explicit JDK 21
selection also works on older checkouts whose Gradle Wrapper cannot run on JDK 25.
In IntelliJ IDEA, set **Gradle JVM** to JDK 21 or 25.

For JDK 25 verification, keep JDK 21 installed and run:

```bash
export JAVA_HOME=/absolute/path/to/jdk-25
export PATH="$JAVA_HOME/bin:$PATH"
./gradlew --version
./gradlew build -PtestJavaVersion=25 --warning-mode=fail
```

Gradle discovers common JDK installation locations. If your installations are
elsewhere, add their absolute paths to `~/.gradle/gradle.properties`:

```properties
org.gradle.java.installations.paths=/absolute/path/to/jdk-21,/absolute/path/to/jdk-25
```

Each test task logs its selected executable and the actual worker JVM version.
CI installs JDK 21 for compilation and runs separate builds and tests on JDK 21
and 25. Changing the test JVM does not change the published bytecode target.

Validate publication generation without credentials or uploading artifacts:

```bash
./gradlew generatePomFileForMavenPublication generateMetadataFileForMavenPublication --warning-mode=fail
```

`build` includes lint and tests. Test reports are in each module's
`build/reports/tests/test/index.html`.

Version choices follow the official [Gradle Java 25 support notes](https://docs.gradle.org/9.1.0/release-notes.html#java-25),
[Kotlin Gradle compatibility table](https://kotlinlang.org/docs/gradle-configure-project.html),
and [ktlint plugin Java 25 support release](https://github.com/JLLeitschuh/ktlint-gradle/releases/tag/v14.0.1).
General library dependency upgrades are handled separately from this build migration.
