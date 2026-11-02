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

## API compatibility and releases

The six published libraries are `core`, `annotation`, `render`, `parser`, `theme`,
and `excel-dsl`. Benchmarks are never published. Each library's `api/*.api` is a
reviewed public JVM ABI baseline (the umbrella module has an empty baseline).
`build` and `checkKotlinAbi` compare compiled APIs with these files. Intentional
changes require review of source, binary and behavioral compatibility before
running `./gradlew updateKotlinAbi` and committing the diff. Never regenerate a
baseline just to make CI pass. Incompatible releases require an explicit migration
plan and version decision; additive APIs also appear in the baseline diff.
ABI checks do not prove source or behavioral compatibility.

CI also compiles and runs the documentation round-trip examples in the standalone
`compatibility` build using Kotlin 2.2.0/JDK 21 and Kotlin 2.3.21/JDK 25.
It resolves only published artifacts, with no project dependency or source sharing.
Keep its examples in sync with `DocumentationTest.kt` when examples change.

To reproduce locally with JDK 21 and 25 installed:

```bash
rm -rf build/verification-repository
./gradlew build publishAllPublicationsToVerificationRepository -Pversion=0.0.0-verification
python3 scripts/verify-publication.py build/verification-repository 0.0.0-verification
./gradlew -p compatibility test -PconsumerKotlin=2.2.0 -PconsumerJava=21 \
  -PartifactVersion=0.0.0-verification -PartifactRepository="$PWD/build/verification-repository"
./gradlew -p compatibility clean test -PconsumerKotlin=2.3.21 -PconsumerJava=25 \
  -PartifactVersion=0.0.0-verification -PartifactRepository="$PWD/build/verification-repository"
```

For a release, first require green CI on the exact commit, review the API diff,
and run the publication checks above. Publish a GitHub release tagged `vX.Y.Z`
(or `X.Y.Z`), or provide `X.Y.Z` explicitly to the manual publish workflow.
Only stable numeric versions are accepted; missing versions, snapshots and malformed
versions fail before building or publishing. The workflow signs and uploads the six
libraries to Maven Central and attests their JARs, including parser. Central and GPG
credentials are required only for public publication. Local verification needs none.
