/** Dad's Workday Gmail -> Supabase importer.
 * Apps Script Settings > Script properties:
 * SUPABASE_URL = https://...supabase.co
 * SUPABASE_SERVICE_ROLE_KEY = your Supabase secret/service_role key
 * ALLOWED_SENDER = your actual work email address
 * Never put secret keys in GitHub or the website.
 * Run setupTrigger() once to schedule sync every 5 minutes.
 */
const LABEL_NAME = 'DadsWorkday';
function setupTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'syncDadWorkday').forEach(ScriptApp.deleteTrigger);
  ScriptApp.newTrigger('syncDadWorkday').timeBased().everyMinutes(5).create();
}
function parseOutlookUtc(value) {
  const raw = String(value || '').trim();
  // Outlook sends UTC timestamps without a timezone suffix (e.g. 2026-11-20T14:00:00.0000000).
  // Treat those as UTC; preserve timestamps already carrying Z or an offset.
  const normalized = /(?:Z|[+-]\d{2}:?\d{2})$/i.test(raw) ? raw : raw + 'Z';
  const date = new Date(normalized);
  if (isNaN(date.getTime())) throw Error('Invalid Outlook UTC timestamp: ' + raw);
  return date;
}
function syncDadWorkday() {
  const props = PropertiesService.getScriptProperties();
  const url = props.getProperty('SUPABASE_URL');
  const key = props.getProperty('SUPABASE_SERVICE_ROLE_KEY');
  const sender = (props.getProperty('ALLOWED_SENDER') || '').trim().toLowerCase();
  if (!url || !key || !sender) throw Error('Missing SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY or ALLOWED_SENDER script property');
  const label = GmailApp.getUserLabelByName(LABEL_NAME);
  if (!label) throw Error('Create Gmail label: ' + LABEL_NAME);
  // Track each message individually: Gmail can thread multiple updates under one subject.
  const threads = GmailApp.search('label:' + LABEL_NAME, 0, 50);
  for (const thread of threads.reverse()) {
    try {
      for (const msg of thread.getMessages()) {
        const doneKey = 'done_' + msg.getId();
        if (props.getProperty(doneKey)) continue;
        const from = msg.getFrom().toLowerCase();
        const match = from.match(/<([^>]+)>/);
        const address = (match ? match[1] : from).trim();
        if (address !== sender) throw Error('Unexpected sender: ' + address);
        if (msg.getSubject().trim() !== 'DadsWorkday Calendar Sync') throw Error('Unexpected subject');
        const body = msg.getPlainBody().replace(/\r/g, '');
        if (!body.includes('DADSWORKDAY_SYNC_V1')) throw Error('Missing sync marker');
        const get = name => {
          const match = body.match(new RegExp('^' + name + ':\\s*(.*)$', 'm'));
          return match ? match[1].trim() : '';
        };
        const id = get('EventID'), action = get('ChangeType').toLowerCase();
        if (!id || id.length > 1000) throw Error('Missing/invalid event ID');
        const endpoint = url.replace(/\/$/, '') + '/rest/v1/family_meetings';
        const headers = {apikey:key,Authorization:'Bearer ' + key};
        let options;
        if (/delete|deleted|remove/.test(action)) {
          options = {method:'delete',headers};
        } else {
          const title = get('Title'), start = parseOutlookUtc(get('Start')), end = parseOutlookUtc(get('End'));
          if (!title || isNaN(start.getTime()) || isNaN(end.getTime()) || end <= start) throw Error('Invalid event data');
          if (end - start > 86400000) throw Error('Event exceeds 24 hours');
          const payload = {outlook_event_id:id,title:title.slice(0,100),starts_at:start.toISOString(),ends_at:end.toISOString()};
          // Do not overwrite status on existing meetings; new ones default to yellow.
          const lookup = UrlFetchApp.fetch(endpoint + '?select=id&outlook_event_id=eq.' + encodeURIComponent(id), {headers,muteHttpExceptions:true});
          if (lookup.getResponseCode() >= 300) throw Error('Lookup failed: ' + lookup.getResponseCode() + ' ' + lookup.getContentText());
          const existing = JSON.parse(lookup.getContentText());
          if (existing.length) {
            options = {method:'patch',headers:{...headers,'Content-Type':'application/json'},payload:JSON.stringify(payload)};
          } else {
            payload.status = 'yellow';
            options = {method:'post',headers:{...headers,'Content-Type':'application/json'},payload:JSON.stringify(payload)};
          }
        }
        options.muteHttpExceptions = true;
        const target = endpoint + (options.method === 'post' ? '' : '?outlook_event_id=eq.' + encodeURIComponent(id));
        const response = UrlFetchApp.fetch(target,options);
        if (response.getResponseCode() >= 300) throw Error('Supabase ' + response.getResponseCode() + ': ' + response.getContentText());
        props.setProperty(doneKey, '1');
      }
    } catch(e) {
      console.error('Thread ' + thread.getId() + ': ' + e.message);
      // Leave unprocessed so it can be retried after fixing the issue.
    }
  }
}

/** One-time correction for already-imported meetings. Preserves status and only
 * updates records that already exist. Run once after replacing the script.
 * Does not alter the processed-message markers or create new meetings.
 */
function repairImportedTimes() {
  const props = PropertiesService.getScriptProperties();
  const url = props.getProperty('SUPABASE_URL');
  const key = props.getProperty('SUPABASE_SERVICE_ROLE_KEY');
  const sender = (props.getProperty('ALLOWED_SENDER') || '').trim().toLowerCase();
  if (!url || !key || !sender) throw Error('Missing script properties');
  const endpoint = url.replace(/\/$/, '') + '/rest/v1/family_meetings';
  const headers = {apikey:key,Authorization:'Bearer ' + key};
  let fixed=0, skipped=0;
  const threads = GmailApp.search('label:' + LABEL_NAME, 0, 500);
  // Oldest first so the latest change for a given event wins.
  const messages = threads.flatMap(t => t.getMessages()).sort((a,b)=>a.getDate()-b.getDate());
  for (const msg of messages) {
    const from = msg.getFrom().toLowerCase();
    const match = from.match(/<([^>]+)>/);
    if ((match ? match[1] : from).trim() !== sender || msg.getSubject().trim() !== 'DadsWorkday Calendar Sync') continue;
    const body=msg.getPlainBody().replace(/\r/g,'');
    if (!body.includes('DADSWORKDAY_SYNC_V1')) continue;
    const get = name => { const m = body.match(new RegExp('^' + name + ':\\s*(.*)$', 'm')); return m ? m[1].trim() : ''; };
    const id=get('EventID');
    if (!id || /delete|remove/i.test(get('ChangeType'))) continue;
    try {
      const start=parseOutlookUtc(get('Start')),end=parseOutlookUtc(get('End'));
      if(end<=start || end-start>86400000) throw Error('Invalid meeting duration');
      const payload={starts_at:start.toISOString(),ends_at:end.toISOString()};
      const response=UrlFetchApp.fetch(endpoint+'?outlook_event_id=eq.'+encodeURIComponent(id),{
        method:'patch',headers:{...headers,'Content-Type':'application/json','Prefer':'return=representation'},
        payload:JSON.stringify(payload),muteHttpExceptions:true
      });
      if(response.getResponseCode()>=300) throw Error('Supabase '+response.getResponseCode()+': '+response.getContentText());
      fixed+=JSON.parse(response.getContentText()).length;
    } catch(e) { skipped++;console.error('Repair failed for message '+msg.getId()+': '+e.message) }
  }
  console.log('Updated '+fixed+' meeting record(s); skipped '+skipped+' message(s) with errors.');
}
