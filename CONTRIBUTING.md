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

## Dependency updates

Shared dependency and plugin versions live in `gradle/libs.versions.toml`.
The standalone consumer imports the same library catalog, but its Kotlin plugin
is deliberately selected by `consumerKotlin` to verify the minimum supported
compiler independently. Keep the Kotlin 2.2.0 consumer CI entry as a compatibility
floor rather than automatically updating it with the producer.

Dependabot checks Gradle builds and GitHub Actions weekly. Minor and patch
updates are grouped; major updates remain separate for review. Require all CI
build and consumer matrix checks in the repository's branch protection/ruleset
before merging updates. The YAML configuration does not itself enforce branch
protection. The active repository ruleset now requires build (21), build (25),
consumer (2.2.0, 21), and consumer (2.3.21, 25), with the branch up to date.
Automatic merging is not enabled.

Kotlin remains at 2.3.21 with a minimum tested consumer of 2.2.0.
Compiler upgrades are compatibility decisions: they can change Kotlin metadata,
stdlib/reflect dependencies and what older compilers can consume. Do not raise
the compiler just to match the latest release. Kotlin updates are excluded from
Dependabot's grouped updates and require independent consumer verification.

Issue #16 upgrades POI to 6.0.0, Kotest to 6.2.6 and the ktlint Gradle plugin to
14.2.0. The ktlint engine remains 1.0.1 to preserve the existing formatting rules.
Gradle's publishing/signing plugins follow the wrapper version.

Official release references checked on 2026-11-02:
[Kotlin](https://kotlinlang.org/docs/releases.html),
[POI changes](https://poi.apache.org/changes.html),
[Kotest](https://github.com/kotest/kotest/releases/tag/v6.2.6),
[ktlint Gradle](https://plugins.gradle.org/plugin/org.jlleitschuh.gradle.ktlint).
Kotlin 2.4.21 is deferred to a dedicated compatibility proposal, including
metadata and transitive stdlib/reflect checks against the supported consumers.
POI 6 removes deprecated APIs and changes DataFormatter's formula-cache default;
the parser reads cell types directly and does not use DataFormatter.
See [POI security guidance](https://poi.apache.org/security.html).
A complete vulnerability audit is outside this dependency migration.

The verified POI runtime graph resolves XMLBeans 5.4.1, Commons Compress 1.28.0,
Commons IO 2.22.0 and Log4j API 2.26.1 (no Log4j Core). Reproduce with
`./gradlew :render:dependencies --configuration runtimeClasspath`.
The minimum consumer test runs with these published dependencies, including
stdlib/reflect 2.3.21; it does not claim compatibility with every older Kotlin
version or every consumer dependency constraint.

Security advisory review on 2026-11-02: the resolved versions are outside the
affected ranges for POI CVE-2025-31672, XMLBeans CVE-2021-23926,
Commons Compress CVE-2024-25710/CVE-2024-26308, Commons IO CVE-2024-47554,
and Log4j API CVE-2026-49844 (fixed in 2.26.1).
Sources: [POI](https://poi.apache.org/security.html),
[XMLBeans](https://xmlbeans.apache.org/status.html),
[Compress](https://commons.apache.org/proper/commons-compress/security.html),
[IO](https://commons.apache.org/proper/commons-io/security.html),
[Log4j](https://logging.apache.org/security.html).
This is a review of those published advisories, not a guarantee against all vulnerabilities.

The ktlint engine 1.8.0 upgrade remains a separate formatting migration because
it adds/promotes rules. JMH plugin 0.7.4 is deferred to a separate benchmark-tooling
change so this POI comparison preserves the 0.7.2 baseline environment.
