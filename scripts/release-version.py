"""Resolve only explicit stable release versions before any publication."""
import re
import sys

event, tag, supplied = sys.argv[1:]
version = tag.removeprefix('v') if event == 'release' else supplied
if not re.fullmatch(r'(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)', version):
    sys.exit('An explicit stable version (e.g. 1.0.0) is required')
print(version)
