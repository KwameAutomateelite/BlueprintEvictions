const b = await this.helpers.getBinaryDataBuffer(0, 'data');
if (b.subarray(0, 2).toString() !== 'PK') throw Error('Saved notice is not a Word file');
return [{json: {docx_base64: b.toString('base64'),
  expected_sha256: require('crypto').createHash('sha256').update(b).digest('hex')}}];
