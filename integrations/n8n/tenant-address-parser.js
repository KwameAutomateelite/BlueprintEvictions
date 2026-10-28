// Parse the premises only. Never borrow a landlord city, state, or ZIP.
function parseConfiguredTenantProperty(raw) {
  const blank = { street: '', city: '', state: '', zip: '', city_state_zip: '' };
  const text = String(raw || '').replace(/\s*[\r\n]+\s*/g, ', ').replace(/\s+/g, ' ').trim();
  if (!text) return blank;
  const tail = text.match(/^(.*?)(?:,\s*|\s+)(CA|California)\.?(?:\s*,\s*|\s*)(\d{5}(?:-\d{4})?)?\s*$/i);
  if (!tail) return { ...blank, street: text };
  const prefix = tail[1].trim().replace(/,+$/, '').trim();
  let street = '', city = '';
  const parts = prefix.split(',').map(s => s.trim()).filter(Boolean);
  if (parts.length >= 2) {
    city = parts.pop();
    street = parts.join(', ');
  } else {
    const cities = ['East Palo Alto','South San Francisco','Los Altos Hills','San Juan Bautista','Mountain View','Morgan Hill','Redwood City','San Francisco','Santa Clara','Santa Cruz','Scotts Valley','San Mateo','San Carlos','San Jose','San José','Los Gatos','Los Altos','San Leandro','San Ramon','Union City','Daly City','Foster City','Half Moon Bay','Menlo Park','Walnut Creek','Sunnyvale','Saratoga','Milpitas','Campbell','Cupertino','Gilroy','Watsonville','Capitola','Aptos','Fremont','Hayward','Oakland','Berkeley','Alameda','Pleasanton','Dublin','Livermore','Newark','Emeryville','Albany','Hollister','Belmont','Millbrae','Burlingame','Atherton','Portola Valley','Woodside','Hillsborough','Brisbane','Colma','Pacifica','Antioch','Concord','Richmond','Pittsburg','Martinez','Danville','Lafayette','Orinda','Monte Sereno','San Bruno'].sort((a,b) => b.length-a.length);
    for (const name of cities) {
      const m = prefix.match(new RegExp('^(.+?)\\s+(' + name + ')$', 'i'));
      if (m) { street = m[1].trim(); city = m[2].trim(); break; }
    }
  }
  // Repeated full locations or an apartment mistaken for a city need review.
  if (!street || !city || !/[A-Za-zÀ-ÿ]/.test(city) || /\d/.test(city)
      || /^(?:apt|apartment|unit|suite|ste|#)\b/i.test(city)
      || /\b(?:CA|California)\s*,?\s*\d{5}\b/i.test(street)) {
    return { ...blank, street: text };
  }
  const zip = tail[3] || '';
  return { street, city, state: 'CA', zip,
    city_state_zip: city + ', CA' + (zip ? ' ' + zip : '') };
}

module.exports = { parseConfiguredTenantProperty };
