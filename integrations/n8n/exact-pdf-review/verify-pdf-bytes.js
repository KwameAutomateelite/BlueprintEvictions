const b = await this.helpers.getBinaryDataBuffer(0, 'data');
if (require('crypto').createHash('sha256').update(b).digest('hex') !==
    $('Render Saved Word as Exact PDF').first().json.pdf_sha256) {
  throw Error('Saved review PDF bytes differ from the checked form. Nothing approved.');
}
return [{json: $('Verify Saved PDF Location').first().json}];
