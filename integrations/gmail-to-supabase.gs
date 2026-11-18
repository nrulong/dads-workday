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
function syncDadWorkday() {
  const props = PropertiesService.getScriptProperties();
  const url = props.getProperty('SUPABASE_URL');
  const key = props.getProperty('SUPABASE_SERVICE_ROLE_KEY');
  const sender = (props.getProperty('ALLOWED_SENDER') || '').trim().toLowerCase();
  if (!url || !key || !sender) throw Error('Missing SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY or ALLOWED_SENDER script property');
  const label = GmailApp.getUserLabelByName(LABEL_NAME);
  if (!label) throw Error('Create Gmail label: ' + LABEL_NAME);
  const processed = GmailApp.getUserLabelByName('DadsWorkdayProcessed') || GmailApp.createLabel('DadsWorkdayProcessed');
  const threads = GmailApp.search('label:' + LABEL_NAME + ' -label:DadsWorkdayProcessed', 0, 50);
  for (const thread of threads.reverse()) {
    try {
      for (const msg of thread.getMessages()) {
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
          const title = get('Title'), start = new Date(get('Start')), end = new Date(get('End'));
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
      }
      thread.addLabel(processed);
    } catch(e) {
      console.error('Thread ' + thread.getId() + ': ' + e.message);
      // Leave unprocessed so it can be retried after fixing the issue.
    }
  }
}
