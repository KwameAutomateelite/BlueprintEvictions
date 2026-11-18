const assert=require('node:assert/strict');
const {mayContainProof,shouldNotifyProofProblem}=require('./proof-notification-policy.cjs');
for(const subject of ['Court login code','Client intake update','New lease','Rent payment received']) {
  assert.equal(mayContainProof({subject,hasAttachments:false,body:{content:'Ordinary correspondence'}}),false);
}
assert.equal(mayContainProof({subject:'Owner v Tenant',hasAttachments:true}),true);
assert.equal(mayContainProof({subject:'Owner v Tenant',hasAttachments:false,body:{content:'Review the POS earlier in this thread'}}),true);
assert.equal(mayContainProof({subject:'Notice was served',hasAttachments:false}),true);
assert.equal(shouldNotifyProofProblem({code:'not_a_pos',silent:true}),false);
assert.equal(shouldNotifyProofProblem({code:'bad_subject'}),false);
assert.equal(shouldNotifyProofProblem({code:'low_confidence'}),true);
assert.equal(shouldNotifyProofProblem({code:'conflicting_case_identity'}),true);
console.log('PASS ordinary correspondence stays quiet; attached/thread proof remains eligible; genuine proof uncertainty remains visible.');
