"""All local Python imports reachable from the service must be in the image."""
import ast
import glob
import shlex
import unittest
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]

class PackagingTests(unittest.TestCase):
    def test_reachable_local_modules_are_copied_into_image(self):
        copied=set()
        for line in (ROOT/'Dockerfile').read_text().splitlines():
            if line.startswith('COPY '):
                for pattern in shlex.split(line)[1:-1]:
                    copied.update(Path(p).resolve() for p in glob.glob(str(ROOT/pattern)))
        pending=[ROOT/'main.py'];seen=set()
        while pending:
            source=pending.pop().resolve()
            if source in seen:continue
            seen.add(source)
            self.assertIn(source,copied,f'{source.name} is imported by the service but absent from the image')
            tree=ast.parse(source.read_text())
            for node in ast.walk(tree):
                modules=[a.name for a in node.names] if isinstance(node,ast.Import) else [node.module] if isinstance(node,ast.ImportFrom) else []
                for module in modules:
                    if not module:continue
                    local=ROOT/(module.split('.')[0]+'.py')
                    if local.is_file():pending.append(local)

if __name__=='__main__':unittest.main()
