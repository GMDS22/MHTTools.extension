import unittest
import sys
import types

# Provide lightweight stubs for pyrevit and Autodesk.Revit imports so tests
# can run outside the Revit/pyRevit environment.
fake_pyrevit = types.ModuleType('pyrevit')
fake_pyrevit.revit = types.SimpleNamespace(doc=None)
def _get_logger():
    import logging
    return logging.getLogger('mht-test')
fake_pyrevit.script = types.SimpleNamespace(get_logger=_get_logger)
sys.modules['pyrevit'] = fake_pyrevit

# Minimal Autodesk.Revit.DB stub
fake_autodesk = types.ModuleType('Autodesk')
fake_revit = types.ModuleType('Autodesk.Revit')
# Provide a minimal DB namespace so `from Autodesk.Revit import DB` works
fake_revit.DB = types.SimpleNamespace()
sys.modules['Autodesk'] = fake_autodesk
sys.modules['Autodesk.Revit'] = fake_revit

# Import the functions under test from the script
from script import apply_template


class TestTokens(unittest.TestCase):
    def test_sys_and_size_tokens_when_classified(self):
        rules = {'COMPANY': 'MHT', 'DISCIPLINE': {'ME': 'ME'}}
        info = {
            '_classified_name': 'Air-Terminal',
            '_classified_system': 'SAD',
            '_classified_size': 'DN150',
            'family_name': 'Air Terminal',
            'types': [{'type_name': 'AT Type', 'params': {}}],
            'instance_params': {},
            'category': 'Mechanical Equipment'
        }
        out = apply_template('<SYS>-<SIZE>-<Type>-<Param:Size>', info, rules)
        # classified system and size should appear
        self.assertIn('SAD', out)
        self.assertIn('DN150', out)

    def test_param_size_fallback_uses_param(self):
        rules = {'COMPANY': 'MHT', 'DISCIPLINE': {'PL': 'PL'}}
        info = {
            'family_name': 'Pipe Fitting',
            'types': [{'type_name': 'Elbow Type', 'params': {'Size': '100'}}],
            'instance_params': {},
            'category': 'Pipe Fittings'
        }
        out = apply_template('<SYS>-<SIZE>-<Type>-<Param:Size>', info, rules)
        # Param:Size should be formatted as DN100
        self.assertIn('DN100', out)
        # Type should contain the cleaned type name
        self.assertIn('Elbow', out)

    def test_air_terminal_round_rect_subtype(self):
        # Load the real naming rules so the classifier has sub_type_keywords
        import json, os
        path = os.path.join(os.getcwd(), 'naming_rules.json')
        with open(path, 'r') as f:
            rules = json.load(f)

        from script import classify_family

        # Make an Air Terminal family whose name contains 'Round' -> expect subtype detection
        info = {
            'family_name': 'Round Diffuser Model X',
            'types': [{'type_name': 'Type A', 'params': {}}],
            'instance_params': {},
            'category': 'Air Terminals'
        }
        # force the generic name to appear used so classifier will attempt more specificity
        used = set(['air-terminal'])
        classified = classify_family(info, rules, used_generic_names=used)
        # classifier should return something that includes the subtype keyword 'round' (case-insensitive)
        self.assertIsNotNone(classified)
        self.assertIn('round', classified.lower())

if __name__ == '__main__':
    unittest.main()
