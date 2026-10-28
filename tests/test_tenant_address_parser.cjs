const assert = require('node:assert/strict');
const test = require('node:test');
const { parseConfiguredTenantProperty: parse } = require('../integrations/n8n/tenant-address-parser.js');

test('comma separated state and ZIP from intake', () => {
  assert.deepEqual(parse('123 EXAMPLE AVENUE, SAN JOSE, CA,95128'), {
    street:'123 EXAMPLE AVENUE',city:'SAN JOSE',state:'CA',zip:'95128',city_state_zip:'SAN JOSE, CA 95128'
  });
});
test('preserves units and handles input whitespace, CRLF and state punctuation', () => {
  for (const value of ['123 Example St, Apt 4, Sunnyvale, CA 94086', '123 Example St, Apt 4\r\nSunnyvale, CA. 94086', '123 Example St, Apt 4, Sunnyvale, California, 94086']) {
    const out = parse(value);
    // A newline separating an apartment from the city must be recognized too.
    assert.equal(out.street, '123 Example St, Apt 4');
    assert.equal(out.city_state_zip, 'Sunnyvale, CA 94086');
  }
});
test('city names are matched longest first without a comma', () => {
  assert.equal(parse('123 Example St East Palo Alto CA 94303').city, 'East Palo Alto');
  assert.equal(parse('123 Example St South San Francisco CA94080').city, 'South San Francisco');
});
test('ZIP+4 and CA comma variants stay on one city line', () => {
  assert.equal(parse('123 Example St, San Jose, CA, 95128-1234').city_state_zip, 'San Jose, CA 95128-1234');
});
test('missing and ambiguous locations remain incomplete instead of borrowing another address', () => {
  for (const value of ['', '123 Example St', '123 Example St, Apt 4, CA 95128', '123 Example St, San Jose, CA 95128, San Jose, CA 95128', '123 Example St, Reno, NV 89501']) {
    const out = parse(value);
    assert.equal(out.city, '', value);
    assert.equal(out.state, '', value);
  }
});
