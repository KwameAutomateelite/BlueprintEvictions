"""Validated rollout configuration; no business record IDs in application logic."""
import json
import os
import re
from pathlib import Path
from dataclasses import dataclass

_RECORD = re.compile(r'rec[A-Za-z0-9]{14}\Z')
_EMAIL = re.compile(r'[^\s@]+@[^\s@]+\.[^\s@]+\Z')

@dataclass(frozen=True)
class SignatureRollout:
    profile: str
    allowed_case_ids: tuple | None
    allowed_signers: tuple | None
    signer_fields: tuple
    test_mode: bool

    def includes(self, record_id):
        return isinstance(record_id, str) and bool(_RECORD.fullmatch(record_id)) and (
            self.allowed_case_ids is None or record_id in self.allowed_case_ids)

    def accepts_signer(self, email, fields):
        signer = str(email or '').strip().lower()
        if not _EMAIL.fullmatch(signer):
            return False
        if self.allowed_signers is not None:
            return signer in self.allowed_signers
        expected = next((str(fields.get(k) or '').strip().lower()
                         for k in self.signer_fields if str(fields.get(k) or '').strip()), '')
        return bool(_EMAIL.fullmatch(expected)) and signer == expected


def parse_rollout(raw):
    d = json.loads(raw)
    if not isinstance(d, dict) or d.get('profile') not in {'kwame-test', 'anne-michelle'}:
        raise ValueError('Unknown signing rollout profile')
    if not {'profile','allowed_case_ids','allowed_signers','signer_fields','test_mode'}.issubset(d):
        raise ValueError('Signing rollout fields must be explicit')
    cases, signers, fields, mode = (d.get(k) for k in ('allowed_case_ids','allowed_signers','signer_fields','test_mode'))
    if cases is not None and (not isinstance(cases,list) or not cases or any(not isinstance(x,str) or not _RECORD.fullmatch(x) for x in cases)):
        raise ValueError('Invalid signing case scope')
    if signers is not None and (not isinstance(signers,list) or not signers or any(not isinstance(x,str) or not _EMAIL.fullmatch(x.strip().lower()) for x in signers)):
        raise ValueError('Invalid signing recipient scope')
    if not isinstance(fields,list) or not fields or any(not isinstance(x,str) or not x.strip() for x in fields) or type(mode) is not bool:
        raise ValueError('Incomplete signing rollout configuration')
    if d['profile']=='kwame-test' and (cases is None or signers is None or mode is not True):
        raise ValueError('Test profile must remain bounded and in test mode')
    if d['profile']=='anne-michelle' and signers is not None:
        raise ValueError('Production signer must match the current case')
    return SignatureRollout(d['profile'],None if cases is None else tuple(cases),None if signers is None else tuple(x.strip().lower() for x in signers),tuple(fields),mode)


def load_rollout():
    raw = os.environ.get('SIGNATURE_ROLLOUT_POLICY')
    if raw is None:
        raw = Path(__file__).with_name('signature-rollout.json').read_text()
    return parse_rollout(raw)
