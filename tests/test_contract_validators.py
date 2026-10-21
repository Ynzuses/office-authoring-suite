"""Synthetic contract regression tests migrated with their owning Skills."""
from __future__ import annotations
import copy
import hashlib
import json
import os
import shutil
import subprocess
import sys
import unittest
import uuid
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterator
SUITE_ROOT = Path(__file__).resolve().parents[1]
VALIDATORS = (('research-document-delivery', 'verify_delivery_record.py'),)

@contextmanager
def fixture_dir() -> Iterator[Path]:
    root = Path(__file__).resolve().parent / f'.contract-{uuid.uuid4().hex}'
    root.mkdir(parents=False, exist_ok=False)
    try:
        yield root
    finally:
        shutil.rmtree(root)

def run_validator(root: Path, skill: str, script: str, data: dict[str, Any], *, output_format: str='json', extra_args: tuple[str, ...]=()) -> subprocess.CompletedProcess[str]:
    state = root / 'record.json'
    state.write_text(json.dumps(data, ensure_ascii=False), encoding='utf-8', errors='strict')
    executable = SUITE_ROOT / skill / 'scripts' / script
    return subprocess.run([sys.executable, str(executable), str(state), '--format', output_format, *extra_args], check=False, capture_output=True, text=True, encoding='utf-8', errors='strict')

def run_validator_bytes(record: Path, skill: str, script: str, source: str) -> subprocess.CompletedProcess[bytes]:
    record.write_text(source, encoding='utf-8', errors='strict')
    executable = SUITE_ROOT / skill / 'scripts' / script
    env = os.environ.copy()
    env['PYTHONIOENCODING'] = 'cp936:strict'
    env['PYTHONUTF8'] = '0'
    return subprocess.run([sys.executable, str(executable), str(record), '--format', 'json'], check=False, capture_output=True, text=False, env=env)

class ContractValidatorTests(unittest.TestCase):

    def test_all_validator_clis_force_utf8_and_reject_non_strict_json(self) -> None:
        with fixture_dir() as root:
            unicode_root = root / '中文 空格'
            unicode_root.mkdir()
            for skill, script in VALIDATORS:
                with self.subTest(skill=skill, case='stdout-utf8'):
                    record = unicode_root / f'{skill}.json'
                    completed = run_validator_bytes(record, skill, script, '{}')
                    stdout = completed.stdout.decode('utf-8', errors='strict')
                    stderr = completed.stderr.decode('utf-8', errors='strict')
                    self.assertEqual(completed.returncode, 1, stderr or stdout)
                    self.assertIn('中文 空格', stdout)
                    self.assertNotIn('Traceback', stderr)
                with self.subTest(skill=skill, case='duplicate-key'):
                    record = unicode_root / f'duplicate-{skill}.json'
                    completed = run_validator_bytes(record, skill, script, '{"重复": 1, "重复": 2}')
                    stderr = completed.stderr.decode('utf-8', errors='strict')
                    self.assertEqual(completed.returncode, 2, stderr)
                    self.assertIn('duplicate object key', stderr)
                    self.assertIn('重复', stderr)
                    self.assertNotIn('Traceback', stderr)
                for constant in ('NaN', 'Infinity', '-Infinity'):
                    with self.subTest(skill=skill, case=constant):
                        record = unicode_root / f'constant-{skill}.json'
                        completed = run_validator_bytes(record, skill, script, f'{{"value": {constant}}}')
                        stderr = completed.stderr.decode('utf-8', errors='strict')
                        self.assertEqual(completed.returncode, 2, stderr)
                        self.assertIn('non-finite number', stderr)
                        self.assertNotIn('Traceback', stderr)

    def test_malformed_collection_members_never_raise_tracebacks(self) -> None:
        malformed: dict[str, dict[str, Any]] = {'research-document-delivery': {'schema_version': '1.0', 'delivery_id': 'delivery:one', 'state': {}, 'content_authority': 'content:one', 'format_authority': 'format:one', 'synchronization_direction': 'single-format', 'artifacts': [], 'delivery_members': [{}], 'validations': [{'gate': 'content', 'artifact_id': {}, 'status': []}], 'invalidated_checks': [], 'unresolved_risks': [], 'next_action': 'test'}}
        script_by_skill = dict(VALIDATORS)
        with fixture_dir() as root:
            for skill, data in malformed.items():
                with self.subTest(skill=skill):
                    completed = run_validator(root, skill, script_by_skill[skill], data)
                    self.assertEqual(completed.returncode, 1, completed.stderr or completed.stdout)
                    self.assertNotIn('Traceback', completed.stderr)

    def test_delivery_record_verifies_current_hash(self) -> None:
        with fixture_dir() as root:
            artifact = root / 'report.md'
            artifact.write_text('synthetic report\n', encoding='utf-8')
            digest = hashlib.sha256(artifact.read_bytes()).hexdigest()
            content_evidence = root / 'content-audit.json'
            content_evidence.write_text('{"content":"passed"}\n', encoding='utf-8')
            content_evidence_hash = hashlib.sha256(content_evidence.read_bytes()).hexdigest()
            package_evidence = root / 'delivery-package.json'
            package_evidence.write_text('{"members":["report.md"]}\n', encoding='utf-8')
            package_evidence_hash = hashlib.sha256(package_evidence.read_bytes()).hexdigest()

            def required_gate(gate: str) -> dict[str, Any]:
                return {'gate': gate, 'evidence_mode': 'file', 'extensions': {}}

            def passed_gate(gate: str, evidence_path: str, evidence_hash: str) -> dict[str, Any]:
                return {'gate': gate, 'status': 'passed', 'artifact_id': 'artifact:report', 'artifact_sha256': digest, 'performed_at': '2026-08-20T12:00:00+08:00', 'evidence': {'kind': 'file', 'path': evidence_path, 'sha256': evidence_hash, 'extensions': {}}, 'not_applicable': None, 'limitations': 'none', 'extensions': {}}
            data = {'schema_version': '2.0', 'delivery_id': 'delivery:one', 'state': 'final', 'content_authority': 'content:one at review.md', 'format_authority': 'format:one at report.md', 'synchronization_direction': 'single-format', 'submitted_baseline_id': None, 'artifacts': [{'artifact_id': 'artifact:report', 'role': 'final', 'path': 'report.md', 'media_type': 'text/markdown', 'sha256': digest, 'frozen': False, 'required_gates': [required_gate('content'), required_gate('delivery-package')], 'extensions': {}}], 'validations': [passed_gate('content', 'content-audit.json', content_evidence_hash), passed_gate('delivery-package', 'delivery-package.json', package_evidence_hash)], 'delivery_members': ['artifact:report'], 'invalidated_checks': [], 'unresolved_risks': [], 'next_action': 'Send after authorization', 'extensions': {}}
            completed = run_validator(root, 'research-document-delivery', 'verify_delivery_record.py', data)
            self.assertEqual(completed.returncode, 0, completed.stderr or completed.stdout)
            payload = json.loads(completed.stdout)
            self.assertEqual(payload['verdict'], 'structural_pass')
            self.assertFalse(payload['delivery_completion_validated'])
            structural_text = run_validator(root, 'research-document-delivery', 'verify_delivery_record.py', data, output_format='text')
            self.assertIn('RESULT: STRUCTURAL PASS', structural_text.stdout)
            self.assertNotIn('RESULT: PASS', structural_text.stdout)
            release_cases: list[tuple[str, dict[str, Any], str]] = []
            missing = copy.deepcopy(data)
            missing['validations'] = missing['validations'][:1]
            release_cases.append(('missing-delivery-package', missing, 'missing_required_gate_validation'))
            unavailable = copy.deepcopy(data)
            unavailable['validations'][0]['status'] = 'unavailable'
            unavailable['validations'][0]['evidence'] = None
            release_cases.append(('unavailable', unavailable, 'unclosed_required_gate'))
            not_applicable = copy.deepcopy(data)
            not_applicable['validations'][1]['status'] = 'not-applicable'
            not_applicable['validations'][1]['evidence'] = None
            not_applicable['validations'][1]['not_applicable'] = {'reason': 'claimed unnecessary despite final handoff', 'authorized_by': 'release owner', 'decided_at': '2026-08-20T12:30:00+08:00', 'extensions': {}}
            release_cases.append(('not-applicable-abuse', not_applicable, 'non_waivable_release_gate'))
            stale = copy.deepcopy(data)
            stale['validations'][0]['artifact_sha256'] = '0' * 64
            release_cases.append(('stale', stale, 'stale_validation'))
            fabricated = copy.deepcopy(data)
            fabricated['validations'][0]['evidence']['path'] = 'invented-audit.json'
            release_cases.append(('fabricated-evidence', fabricated, 'missing_evidence_file'))
            old_schema = copy.deepcopy(data)
            old_schema['schema_version'] = '1.0'
            release_cases.append(('old-schema', old_schema, 'schema_version_migration'))
            for case, invalid, expected_code in release_cases:
                with self.subTest(case=case):
                    completed = run_validator(root, 'research-document-delivery', 'verify_delivery_record.py', invalid)
                    self.assertEqual(completed.returncode, 1, completed.stderr or completed.stdout)
                    self.assertIn(expected_code, completed.stdout)
            candidate = copy.deepcopy(data)
            candidate['state'] = 'candidate'
            candidate['validations'][0]['status'] = 'unavailable'
            candidate['validations'][0]['evidence'] = {'kind': 'manual', 'description': 'content review awaits the domain owner', 'source': 'candidate handoff note', 'extensions': {}}
            candidate['validations'][0]['not_applicable'] = None
            candidate_result = run_validator(root, 'research-document-delivery', 'verify_delivery_record.py', candidate)
            self.assertEqual(candidate_result.returncode, 0, candidate_result.stderr or candidate_result.stdout)
            self.assertEqual(json.loads(candidate_result.stdout)['verdict'], 'structural_pass')

    def test_delivery_record_docx_release_requires_full_gate_closure(self) -> None:
        with fixture_dir() as root:
            artifact = root / 'report.docx'
            artifact.write_bytes(b'synthetic-docx-placeholder')
            digest = hashlib.sha256(artifact.read_bytes()).hexdigest()
            evidence = root / 'content-audit.json'
            evidence.write_text('{"content":"passed"}\n', encoding='utf-8')
            evidence_hash = hashlib.sha256(evidence.read_bytes()).hexdigest()
            data = {'schema_version': '2.0', 'delivery_id': 'delivery:docx', 'state': 'final', 'content_authority': 'content:docx', 'format_authority': 'format:docx', 'synchronization_direction': 'single-format', 'submitted_baseline_id': None, 'artifacts': [{'artifact_id': 'artifact:report', 'role': 'final', 'path': 'report.docx', 'media_type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'sha256': digest, 'frozen': False, 'required_gates': [{'gate': 'content', 'evidence_mode': 'file', 'extensions': {}}], 'extensions': {}}], 'validations': [{'gate': 'content', 'status': 'passed', 'artifact_id': 'artifact:report', 'artifact_sha256': digest, 'performed_at': '2026-08-20T12:00:00+08:00', 'evidence': {'kind': 'file', 'path': 'content-audit.json', 'sha256': evidence_hash, 'extensions': {}}, 'not_applicable': None, 'limitations': 'none', 'extensions': {}}], 'delivery_members': ['artifact:report'], 'invalidated_checks': [], 'unresolved_risks': [], 'next_action': 'Complete package, render, and delivery-package gates', 'extensions': {}}
            completed = run_validator(root, 'research-document-delivery', 'verify_delivery_record.py', data)
            self.assertEqual(completed.returncode, 1)
            self.assertIn('missing_required_gate_contract', completed.stdout)
            self.assertIn('rendered', completed.stdout)
            self.assertIn('delivery-package', completed.stdout)
            fabricated_render = copy.deepcopy(data)
            fabricated_render['artifacts'][0]['required_gates'].extend([{'gate': 'package', 'evidence_mode': 'file', 'extensions': {}}, {'gate': 'rendered', 'evidence_mode': 'file', 'extensions': {}}, {'gate': 'delivery-package', 'evidence_mode': 'file', 'extensions': {}}])
            for gate in ('package', 'delivery-package'):
                fabricated_render['validations'].append({**copy.deepcopy(fabricated_render['validations'][0]), 'gate': gate})
            fabricated_render['validations'].append({**copy.deepcopy(fabricated_render['validations'][0]), 'gate': 'rendered', 'evidence': {'kind': 'file', 'path': 'fictional-render-report.json', 'sha256': '0' * 64, 'extensions': {}}})
            render_result = run_validator(root, 'research-document-delivery', 'verify_delivery_record.py', fabricated_render)
            self.assertEqual(render_result.returncode, 1)
            self.assertIn('missing_evidence_file', render_result.stdout)
            manual_render = copy.deepcopy(fabricated_render)
            manual_render['validations'][-1]['evidence'] = {'kind': 'manual', 'description': 'claimed visual inspection without a captured report', 'source': 'operator assertion', 'extensions': {}}
            manual_result = run_validator(root, 'research-document-delivery', 'verify_delivery_record.py', manual_render)
            self.assertEqual(manual_result.returncode, 1)
            self.assertIn('release_requires_file_evidence', manual_result.stdout)
            waived_render = copy.deepcopy(fabricated_render)
            waived_render['validations'][-1].update({'status': 'not-applicable', 'evidence': None, 'not_applicable': {'reason': 'authoritative renderer was not available', 'authorized_by': 'release owner', 'decided_at': '2026-08-20T13:00:00+08:00', 'extensions': {}}})
            waived_result = run_validator(root, 'research-document-delivery', 'verify_delivery_record.py', waived_render)
            self.assertEqual(waived_result.returncode, 1)
            self.assertIn('non_waivable_release_gate', waived_result.stdout)
if __name__ == '__main__':
    unittest.main()
