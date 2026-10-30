const reviewer = $('Validate Review Context').first().json.reviewer;
const permissions = $input.first().json.value || [];
const readable = permissions.some(p => {
  const emails = [p.invitation?.email, p.grantedToV2?.user?.email,
    ...(p.grantedToIdentitiesV2 || []).map(x => x.user?.email)].filter(Boolean).map(x => x.toLowerCase());
  return emails.includes(reviewer) && p.roles?.some(x => ['read', 'write', 'owner'].includes(x));
});
if (!readable) throw Error('The reviewer cannot open the exact PDF. Nothing approved.');
return [{json: {readable: true}}];
