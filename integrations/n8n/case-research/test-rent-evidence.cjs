const {strict:assert}=require('node:assert');
const {rentEvidence}=require('./rent-evidence.cjs');
const base={value:'November 2026 rent: $4,300.00\nTotal: $4,300.00',current:"November' rent $4300",
 quote:'November rent is late',support:[],orderDate:'2026-11-16',
 source:{source_id:'orders:case-mail',source_kind:'email',date:'2026-11-11T00:36:44Z',text:'November rent is late\nOn Nov 10, 2026, someone wrote: old thread'}};
assert.equal(rentEvidence(base).period_basis.year,2026);
for(const change of [
 {value:'November 2025 rent: $4,300.00'},
 {current:"November' rent $4300",value:'November 2026 rent: $4,500.00'},
 {current:'October rent $4300'},
 {orderDate:'2027-11-16'},
 {source:{...base.source,date:'2026-10-11T00:36:44Z'}},
 {source:{...base.source,text:'Nothing owed\nOn Nov 10, 2026, someone wrote: November rent is late'}},
 {quote:'Last year November rent is late',source:{...base.source,text:'Last year November rent is late'}},
 {value:'November 2026 rent: $43,000.00'},
 {value:'November 2026 rent: $4,300.00\nDecember 2026 rent: $4,300.00'}
])assert(rentEvidence({...base,...change}).error,JSON.stringify(change));
// Currency formatting must not reject equivalent literal amounts.
const explicit={...base,current:'',quote:'November 2026 unpaid rent is $4300',support:[]};
assert(!rentEvidence(explicit).error);
assert(rentEvidence({...explicit,value:'November 2026 rent: $430.00'}).error);
assert(rentEvidence({...base,current:'',support:[{quote:'Lease signed 2026: monthly rent $4300'}]}).error);
console.log('13 rent-evidence checks passed: current-period evidence, amount equality, stale/quoted mail, wrong year, added month and rate-only rejection');
