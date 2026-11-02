"""Verify the complete temporary Maven repository, including transitive metadata."""
import json
import sys
from pathlib import Path
import xml.etree.ElementTree as ET

root = Path(sys.argv[1]) / 'io/clroot/excel'
version = sys.argv[2]
modules = {'core', 'annotation', 'render', 'parser', 'theme', 'excel-dsl'}
assert {p.name for p in root.iterdir()} == modules, 'Unexpected publication modules'
ns = {'m': 'http://maven.apache.org/POM/4.0.0'}
for module in sorted(modules):
    base = root / module / version / f'{module}-{version}'
    for suffix in ('.jar', '-sources.jar', '-javadoc.jar', '.pom', '.module'):
        assert Path(str(base) + suffix).is_file(), f'Missing {base}{suffix}'
    pom = ET.parse(str(base) + '.pom')
    assert pom.findtext('m:version', namespaces=ns) == version
    pom_deps = set()
    for dep in pom.findall('m:dependencies/m:dependency', ns):
        if dep.findtext('m:groupId', namespaces=ns) == 'io.clroot.excel':
            name = dep.findtext('m:artifactId', namespaces=ns)
            assert name in modules
            assert dep.findtext('m:version', namespaces=ns) == version
            pom_deps.add(name)
    metadata = json.loads(Path(str(base) + '.module').read_text())
    assert metadata['component']['version'] == version
    module_deps = set()
    for variant in metadata['variants']:
        for dep in variant.get('dependencies', []):
            if dep['group'] == 'io.clroot.excel':
                assert dep['module'] in modules
                assert dep['version']['requires'] == version
                module_deps.add(dep['module'])
    assert pom_deps == module_deps, f'Dependency mismatch: {module}'
    if module == 'excel-dsl':
        assert pom_deps == modules - {'excel-dsl'}
print('Verified six publications and their POM/module metadata')
