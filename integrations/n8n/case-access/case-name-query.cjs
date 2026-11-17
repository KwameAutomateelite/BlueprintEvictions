function prepareCaseQuery($json) {
const query=String($json.query||'').normalize('NFKC').trim();if(query.length<2||query.length>160)throw Error('A case name, person name or property address is required');const stop=new Set(['v','vs','versus','show','me','all','notices','notice','retainers','retainer','for','the','case']);const terms=[...new Set(query.toLowerCase().replace(/[^a-z0-9]+/g,' ').split(' ').filter(x=>x.length>=2&&!stop.has(x)))].slice(0,12);if(!terms.length)throw Error('Provide the actual case or party name');const hay='LOWER({Case Name}&" "&{Tenant\'s Property Address})';const clauses=terms.map(t=>'SEARCH('+JSON.stringify(t)+','+hay+')>0');return [{json:{query,terms,formula:'AND('+clauses.join(',')+')'}}];
}
module.exports={prepareCaseQuery};
