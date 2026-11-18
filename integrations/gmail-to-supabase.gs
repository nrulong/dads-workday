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
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    const props = PropertiesService.getScriptProperties();
    const url = props.getProperty('SUPABASE_URL');
    const key = props.getProperty('SUPABASE_SERVICE_ROLE_KEY');
    const sender = (props.getProperty('ALLOWED_SENDER') || '').trim().toLowerCase();
    if (!url || !key || !sender) throw Error('Missing script properties');
    const endpoint = url.replace(/\/$/, '') + '/rest/v1/family_meetings';
    const headers = {apikey:key,Authorization:'Bearer '+key,'Content-Type':'application/json'};
    const messages = GmailApp.search('label:DadsWorkday subject:"DadsWorkday Calendar Batch" newer_than:7d',0,50)
      .flatMap(t=>t.getMessages()).filter(m=>{
        const from=m.getFrom().toLowerCase(), match=from.match(/<([^>]+)>/);
        return (match?match[1]:from).trim()===sender && m.getSubject().trim()==='DadsWorkday Calendar Batch';
      }).sort((a,b)=>b.getDate()-a.getDate());
    const msg=messages[0];
    if (!msg || props.getProperty('last_batch_id')===msg.getId()) return;
    const raw=msg.getPlainBody().trim();
    const events=JSON.parse(raw);
    if (!Array.isArray(events)) throw Error('Expected a JSON array');
    const seen=new Set();
    const rows=events.map(e=>{
      const id=String(e.eventID||'').trim();
      const start=parseOutlookUtc(e.start),end=parseOutlookUtc(e.end);
      if(!id || id.length>1000 || seen.has(id) || end<=start) throw Error('Invalid or duplicate event');
      seen.add(id);
      return {outlook_event_id:id,title:String(e.title||'Untitled meeting').slice(0,100),starts_at:start.toISOString(),ends_at:end.toISOString()};
    });
    function request(target,options){
      const response=UrlFetchApp.fetch(target,{headers,muteHttpExceptions:true,...options});
      if(response.getResponseCode()>=300) throw Error('Supabase '+response.getResponseCode()+': '+response.getContentText());
      return response.getContentText();
    }
    let changed=0;
    for(const row of rows){
      const target=endpoint+'?outlook_event_id=eq.'+encodeURIComponent(row.outlook_event_id);
      const existing=JSON.parse(request(target+'&select=id,title,starts_at,ends_at',{}))[0];
      if(existing){
        if(existing.title===row.title && new Date(existing.starts_at).getTime()===new Date(row.starts_at).getTime() && new Date(existing.ends_at).getTime()===new Date(row.ends_at).getTime()) continue;
        request(target,{method:'patch',payload:JSON.stringify(row)});
      } else {
        request(endpoint,{method:'post',payload:JSON.stringify({...row,status:'yellow'})});
      }
      changed++;
    }
    // Do not delete absent events until snapshot boundaries and pagination are verified.
    props.setProperty('last_batch_id',msg.getId());
    console.log('Batch received: '+rows.length+' meetings; '+changed+' added or updated. No deletions.');
  } finally {lock.releaseLock();}
}
