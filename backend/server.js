import http from 'node:http';
import crypto from 'node:crypto';
import {readFile} from 'node:fs/promises';
import {fileURLToPath} from 'node:url';
import path from 'node:path';
import pg from 'pg';

const {Pool}=pg;
const PORT=Number(process.env.PORT||3000);
const DATABASE_URL=process.env.DATABASE_URL;
const SESSION_SECRET=process.env.SESSION_SECRET;
const ADMIN_SETUP_KEY=process.env.ADMIN_SETUP_KEY;
if(!DATABASE_URL||!SESSION_SECRET||SESSION_SECRET.length<32||!ADMIN_SETUP_KEY||ADMIN_SETUP_KEY.length<20){
  console.error('Set DATABASE_URL, SESSION_SECRET (32+ chars), and ADMIN_SETUP_KEY (20+ chars) in Railway variables.');process.exit(1);
}
const pool=new Pool({connectionString:DATABASE_URL});
await pool.query(`CREATE TABLE IF NOT EXISTS users(
 id BIGSERIAL PRIMARY KEY, username TEXT UNIQUE NOT NULL, password_hash TEXT NOT NULL,
 role TEXT NOT NULL CHECK(role IN ('admin','viewer')), created_at TIMESTAMPTZ DEFAULT now());
 CREATE TABLE IF NOT EXISTS sessions(
 token_hash TEXT PRIMARY KEY, user_id BIGINT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
 expires_at TIMESTAMPTZ NOT NULL);
 CREATE TABLE IF NOT EXISTS meetings(
 id BIGSERIAL PRIMARY KEY, title TEXT NOT NULL, starts_at TIMESTAMPTZ NOT NULL,
 ends_at TIMESTAMPTZ NOT NULL, status TEXT NOT NULL CHECK(status IN ('green','yellow','red')),
 created_at TIMESTAMPTZ DEFAULT now());
 CREATE INDEX IF NOT EXISTS meetings_start_idx ON meetings(starts_at);`);
const publicDir=path.join(path.dirname(fileURLToPath(import.meta.url)),'public');
const hashPassword=p=>{const salt=crypto.randomBytes(16).toString('hex');return salt+':'+crypto.scryptSync(p,salt,64).toString('hex')};
const checkPassword=(p,stored)=>{try{const [salt,digest]=stored.split(':');const expected=Buffer.from(digest,'hex');const got=crypto.scryptSync(p,salt,expected.length);return expected.length===got.length&&crypto.timingSafeEqual(expected,got)}catch{return false}};
const sha=s=>crypto.createHash('sha256').update(s).digest('hex');
const cookie=(name,value,age)=>name+'='+encodeURIComponent(value)+'; Path=/; HttpOnly; Secure; SameSite=Strict; Max-Age='+age;
const reply=(res,status,data,headers={})=>{res.writeHead(status,{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store','X-Content-Type-Options':'nosniff',...headers});res.end(JSON.stringify(data))};
const fail=(res,status,message)=>reply(res,status,{error:message});
const readBody=async req=>{let raw='';for await(const chunk of req){raw+=chunk;if(raw.length>16000)throw Error('Request too large')}return JSON.parse(raw||'{}')};
const attempts=new Map();
function limited(key){const now=Date.now();let a=attempts.get(key)||[];a=a.filter(t=>now-t<15*60*1000);if(a.length>=10)return true;a.push(now);attempts.set(key,a);if(attempts.size>2000)attempts.clear();return false}
async function currentUser(req){const token=(req.headers.cookie||'').split(';').map(x=>x.trim()).find(x=>x.startsWith('dwd_session='))?.split('=')[1];if(!token)return null;const r=await pool.query('SELECT u.id,u.username,u.role FROM sessions s JOIN users u ON u.id=s.user_id WHERE s.token_hash=$1 AND s.expires_at>now()',[sha(decodeURIComponent(token))]);return r.rows[0]||null}
const validUser=s=>typeof s==='string'&&/^[a-zA-Z0-9._-]{3,40}$/.test(s);
const validPass=s=>typeof s==='string'&&s.length>=12&&s.length<=200;
const validDate=s=>typeof s==='string'&&/^\d{4}-\d\d-\d\dT/.test(s)&&!Number.isNaN(Date.parse(s));
const validMeeting=m=>typeof m.title==='string'&&m.title.trim().length>0&&m.title.length<=100&&validDate(m.start)&&validDate(m.end)&&Date.parse(m.end)>Date.parse(m.start)&&Date.parse(m.end)-Date.parse(m.start)<=24*3600000&&['green','yellow','red'].includes(m.status);
const server=http.createServer(async(req,res)=>{try{
 const u=new URL(req.url,'http://localhost');const route=u.pathname;const method=req.method;
 res.setHeader('X-Frame-Options','DENY');res.setHeader('Referrer-Policy','no-referrer');res.setHeader('Content-Security-Policy',"default-src 'self'; style-src 'self'; script-src 'self'; connect-src 'self'; base-uri 'none'; frame-ancestors 'none'");
 if(route==='/health')return reply(res,200,{ok:true});
 if(method==='GET'&&(route==='/'||route==='/app.js'||route==='/style.css')){const name=route==='/'?'index.html':route.slice(1);const types={'index.html':'text/html; charset=utf-8','app.js':'text/javascript; charset=utf-8','style.css':'text/css; charset=utf-8'};res.writeHead(200,{'Content-Type':types[name],'Cache-Control':'no-store'});return res.end(await readFile(path.join(publicDir,name)))}
 if(!route.startsWith('/api/'))return fail(res,404,'Not found');
 if(method!=='GET'){const origin=req.headers.origin;const host=req.headers.host;if(origin&&new URL(origin).host!==host)return fail(res,403,'Origin mismatch')}
 if(route==='/api/setup-status'&&method==='GET'){const n=await pool.query('SELECT COUNT(*)::int AS n FROM users');return reply(res,200,{needsSetup:n.rows[0].n===0})}
 if(route==='/api/setup'&&method==='POST'){if(limited('setup:'+req.socket.remoteAddress))return fail(res,429,'Try later');const b=await readBody(req);if(!validUser(b.username)||!validPass(b.password)||typeof b.setupKey!=='string'||b.setupKey!==ADMIN_SETUP_KEY)return fail(res,400,'Invalid setup details');const n=await pool.query('SELECT COUNT(*)::int AS n FROM users');if(n.rows[0].n!==0)return fail(res,403,'Setup already complete');await pool.query("INSERT INTO users(username,password_hash,role) VALUES($1,$2,'admin')",[b.username,hashPassword(b.password)]);return reply(res,201,{ok:true})}
 if(route==='/api/login'&&method==='POST'){if(limited('login:'+req.socket.remoteAddress))return fail(res,429,'Too many attempts; try later');const b=await readBody(req);const r=await pool.query('SELECT * FROM users WHERE username=$1',[String(b.username||'')]);if(!r.rows[0]||!checkPassword(String(b.password||''),r.rows[0].password_hash))return fail(res,401,'Incorrect username or password');const token=crypto.randomBytes(32).toString('base64url');await pool.query("INSERT INTO sessions(token_hash,user_id,expires_at) VALUES($1,$2,now()+interval '14 days')",[sha(token),r.rows[0].id]);return reply(res,200,{ok:true},{'Set-Cookie':cookie('dwd_session',token,14*86400)})}
 const user=await currentUser(req);if(!user)return fail(res,401,'Please sign in');
 if(route==='/api/me'&&method==='GET')return reply(res,200,{username:user.username,role:user.role});
 if(route==='/api/logout'&&method==='POST'){const token=(req.headers.cookie||'').split(';').map(x=>x.trim()).find(x=>x.startsWith('dwd_session='))?.split('=')[1];if(token)await pool.query('DELETE FROM sessions WHERE token_hash=$1',[sha(decodeURIComponent(token))]);return reply(res,200,{ok:true},{'Set-Cookie':cookie('dwd_session','',0)})}
 if(route==='/api/meetings'&&method==='GET'){const from=u.searchParams.get('from'),to=u.searchParams.get('to');if(!validDate(from)||!validDate(to)||Date.parse(to)-Date.parse(from)>31*86400000)return fail(res,400,'Invalid date range');const r=await pool.query('SELECT id::text,title,starts_at AS start,ends_at AS "end",status FROM meetings WHERE starts_at<$2 AND ends_at>$1 ORDER BY starts_at LIMIT 500',[from,to]);return reply(res,200,{meetings:r.rows})}
 if(user.role!=='admin')return fail(res,403,'Admin access required');
 if(route==='/api/users'&&method==='POST'){const b=await readBody(req);if(!validUser(b.username)||!validPass(b.password))return fail(res,400,'Username or password invalid');try{await pool.query("INSERT INTO users(username,password_hash,role) VALUES($1,$2,'viewer')",[b.username,hashPassword(b.password)]);return reply(res,201,{ok:true})}catch(e){if(e.code==='23505')return fail(res,409,'Username already exists');throw e}}
 if(route==='/api/meetings'&&method==='POST'){const b=await readBody(req);if(!validMeeting(b))return fail(res,400,'Invalid meeting');const r=await pool.query('INSERT INTO meetings(title,starts_at,ends_at,status) VALUES($1,$2,$3,$4) RETURNING id::text',[b.title.trim(),b.start,b.end,b.status]);return reply(res,201,{id:r.rows[0].id})}
 const match=route.match(/^\/api\/meetings\/(\d+)$/);
 if(match&&method==='PATCH'){const b=await readBody(req);if(!['green','yellow','red'].includes(b.status))return fail(res,400,'Invalid status');const r=await pool.query('UPDATE meetings SET status=$1 WHERE id=$2 RETURNING id',[b.status,match[1]]);return r.rowCount?reply(res,200,{ok:true}):fail(res,404,'Not found')}
 if(match&&method==='DELETE'){await pool.query('DELETE FROM meetings WHERE id=$1',[match[1]]);return reply(res,200,{ok:true})}
 return fail(res,404,'Not found');
 }catch(e){console.error(e.message);return fail(res,e.message==='Request too large'?413:500,'Server error')}});
server.listen(PORT,'0.0.0.0',()=>console.log('Dad\'s Workday listening on '+PORT));
