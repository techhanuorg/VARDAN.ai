require('dotenv').config();
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const express = require('express');
const multer = require('multer');
const cron = require('node-cron');
const XLSX = require('xlsx');
const mammoth = require('mammoth');
const pdf = require('pdf-parse');
const Groq = require('groq-sdk');
const { DatabaseSync } = require('node:sqlite');

const app = express();
const port = process.env.PORT || 3000;
const root = __dirname;
const dataDir = path.join(root, 'data');
const uploadDir = path.join(root, 'uploads');
fs.mkdirSync(dataDir, { recursive: true }); fs.mkdirSync(uploadDir, { recursive: true });
const db = new DatabaseSync(path.join(dataDir, 'vardan.sqlite'));
db.exec(`CREATE TABLE IF NOT EXISTS state (id INTEGER PRIMARY KEY CHECK(id=1), payload TEXT NOT NULL, updated_at TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS files (id TEXT PRIMARY KEY, name TEXT NOT NULL, type TEXT, category TEXT NOT NULL, path TEXT NOT NULL, extracted_text TEXT, created_at TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS deliveries (id TEXT PRIMARY KEY, campaign_id TEXT, recipient TEXT, status TEXT NOT NULL, error TEXT, created_at TEXT NOT NULL);`);
const emptyState = { hospital:{}, doctors:[], contacts:[], groups:[], campaigns:[], media:[], knowledge:[], broadcasts:[] };
function getState(){ const row=db.prepare('SELECT payload FROM state WHERE id=1').get(); return row ? JSON.parse(row.payload) : structuredClone(emptyState); }
function putState(body){ const payload=JSON.stringify({ ...emptyState, ...body }); db.prepare('INSERT INTO state(id,payload,updated_at) VALUES(1,?,?) ON CONFLICT(id) DO UPDATE SET payload=excluded.payload,updated_at=excluded.updated_at').run(payload,new Date().toISOString()); return JSON.parse(payload); }
function id(){ return crypto.randomUUID(); }
const upload=multer({storage:multer.diskStorage({destination:uploadDir,filename:(_,f,cb)=>cb(null,`${Date.now()}-${crypto.randomUUID()}${path.extname(f.originalname)}`)}),limits:{fileSize:20*1024*1024}});
app.use(express.json({limit:'15mb'})); app.use('/uploads',express.static(uploadDir)); app.use(express.static(root));
app.get('/api/health',(_,res)=>res.json({ok:true,service:'VARDAN.ai',integrations:{groq:!!process.env.GROQ_API_KEY,whatsapp:!!(process.env.WHATSAPP_TOKEN&&process.env.WHATSAPP_PHONE_NUMBER_ID),sheets:!!process.env.GOOGLE_SHEETS_WEBHOOK_URL}}));
app.get('/api/state',(_,res)=>res.json(getState()));
app.put('/api/state',(req,res)=>res.json(putState(req.body)));

function normalizePhone(v=''){return String(v).replace(/[^\d+]/g,'');}
function contactFromRow(row={}){ const get=(...keys)=>{const k=Object.keys(row).find(x=>keys.some(key=>x.toLowerCase().includes(key)));return k?String(row[k]||''):''}; const phone=normalizePhone(get('phone','mobile','tel','number')); return {id:id(),name:get('name','first','full name'),phone,email:get('email'),company:get('company','organization','org'),notes:get('note','comment'),valid:/^\+?\d{7,15}$/.test(phone)}; }
function parseVcf(text){return text.split(/END:VCARD/i).map(v=>({name:(v.match(/(?:FN):(.+)/i)||[])[1]||'',phone:(v.match(/TEL[^:]*:(.+)/i)||[])[1]||'',email:(v.match(/EMAIL[^:]*:(.+)/i)||[])[1]||'',company:(v.match(/ORG:(.+)/i)||[])[1]||'',notes:(v.match(/NOTE:(.+)/i)||[])[1]||''})).filter(c=>c.name||c.phone).map(c=>({...c,id:id(),phone:normalizePhone(c.phone),valid:/^\+?\d{7,15}$/.test(normalizePhone(c.phone))}));}
app.post('/api/contacts/preview',upload.single('file'),(req,res)=>{try{const f=req.file;if(!f)return res.status(400).json({error:'A file is required'});let rows=[];if(/\.vcf$/i.test(f.originalname))rows=parseVcf(fs.readFileSync(f.path,'utf8'));else {const wb=XLSX.readFile(f.path);rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''}).map(contactFromRow);}fs.unlinkSync(f.path);const seen=new Set();rows=rows.filter(c=>{const key=c.phone||`${c.name}|${c.email}`;if(seen.has(key))return false;seen.add(key);return key.trim();});res.json({contacts:rows});}catch(e){res.status(400).json({error:`Could not read contacts: ${e.message}`})}});
app.post('/api/contacts/import',(req,res)=>{const state=getState();const known=new Set(state.contacts.map(c=>normalizePhone(c.phone)));const added=(req.body.contacts||[]).filter(c=>{const p=normalizePhone(c.phone);if(!p||known.has(p))return false;known.add(p);return true;}).map(c=>({...c,id:c.id||id(),phone:normalizePhone(c.phone),valid:/^\+?\d{7,15}$/.test(normalizePhone(c.phone))}));state.contacts.push(...added);putState(state);mirrorContacts(added).catch(console.error);res.json({added:added.length,contacts:added});});
async function mirrorContacts(contacts){if(!process.env.GOOGLE_SHEETS_WEBHOOK_URL||!contacts.length)return;await fetch(process.env.GOOGLE_SHEETS_WEBHOOK_URL,{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({contacts})});}

async function extractText(file){if(/\.txt$/i.test(file.originalname))return fs.readFileSync(file.path,'utf8');if(/\.docx$/i.test(file.originalname))return (await mammoth.extractRawText({path:file.path})).value;if(/\.pdf$/i.test(file.originalname))return (await pdf(fs.readFileSync(file.path))).text;return '';}
app.post('/api/files/:category',upload.single('file'),async(req,res)=>{try{const category=req.params.category;if(!['media','knowledge'].includes(category)||!req.file)return res.status(400).json({error:'Choose a media or knowledge file'});const item={id:id(),name:req.file.originalname,type:req.file.mimetype||path.extname(req.file.originalname),category,path:`/uploads/${req.file.filename}`,createdAt:new Date().toISOString()};const extracted=category==='knowledge'?await extractText(req.file):'';db.prepare('INSERT INTO files VALUES(?,?,?,?,?,?,?)').run(item.id,item.name,item.type,item.category,item.path,extracted,item.createdAt);const state=getState();state[category].push(item);putState(state);res.json(item);}catch(e){res.status(400).json({error:`Upload failed: ${e.message}`})}});

async function sendWhatsApp(phone,text){if(!process.env.WHATSAPP_TOKEN||!process.env.WHATSAPP_PHONE_NUMBER_ID)throw new Error('WhatsApp Cloud API is not configured');const r=await fetch(`https://graph.facebook.com/v20.0/${process.env.WHATSAPP_PHONE_NUMBER_ID}/messages`,{method:'POST',headers:{Authorization:`Bearer ${process.env.WHATSAPP_TOKEN}`,'content-type':'application/json'},body:JSON.stringify({messaging_product:'whatsapp',to:phone.replace(/^\+/,''),type:'text',text:{body:text}})});if(!r.ok)throw new Error(await r.text());return r.json();}
function personalized(text,contact,doctor={}){return text.replace(/{{\s*patient_name\s*}}/gi,contact.name||'Patient').replace(/{{\s*doctor_name\s*}}/gi,doctor.name||'Doctor').replace(/{{\s*language\s*}}/gi,contact.language||'English');}
app.post('/api/broadcasts/send',async(req,res)=>{const state=getState(),{message,audience='All patients'}=req.body;let recipients=state.contacts;if(audience.startsWith('Group:'))recipients=recipients.filter(c=>c.group===audience.slice(7));const results=[];for(const c of recipients.filter(c=>c.valid)){try{await sendWhatsApp(c.phone,personalized(message,c));results.push({recipient:c.phone,status:'sent'});}catch(e){results.push({recipient:c.phone,status:'failed',error:e.message});} }for(const r of results)db.prepare('INSERT INTO deliveries VALUES(?,?,?,?,?,?)').run(id(),null,r.recipient,r.status,r.error||null,new Date().toISOString());const log={id:id(),message,audience,total:recipients.length,success:results.filter(r=>r.status==='sent').length,failed:results.filter(r=>r.status==='failed').length,date:new Date().toLocaleDateString()};state.broadcasts.unshift(log);putState(state);res.json(log);});
app.get('/api/deliveries',(_,res)=>res.json(db.prepare('SELECT * FROM deliveries ORDER BY created_at DESC LIMIT 200').all()));
app.post('/api/ai/ask',async(req,res)=>{if(!process.env.GROQ_API_KEY)return res.status(503).json({error:'Add GROQ_API_KEY to .env to activate AI.'});const question=String(req.body.question||'').slice(0,4000),state=getState();const docs=db.prepare("SELECT name,extracted_text FROM files WHERE category='knowledge' ORDER BY created_at DESC").all().map(x=>`SOURCE: ${x.name}\n${x.extracted_text||''}`).join('\n\n').slice(0,24000);const doctors=state.doctors.map(d=>`${d.name}: ${d.specialization}; ${d.availableDays}; ${d.slots}; fees ${d.fees}`).join('\n');try{const groq=new Groq({apiKey:process.env.GROQ_API_KEY});const out=await groq.chat.completions.create({model:process.env.GROQ_MODEL||'llama-3.3-70b-versatile',messages:[{role:'system',content:`You are the hospital's patient assistant. Only use provided hospital context. If unavailable, say the owner has not added this information. Never diagnose or prescribe. Doctors:\n${doctors}\n\nKnowledge:\n${docs}`},{role:'user',content:question}],temperature:.2,max_tokens:700});res.json({answer:out.choices[0].message.content});}catch(e){res.status(502).json({error:`AI request failed: ${e.message}`})}});
// Checks scheduled campaigns each minute. Configure a cron expression in a campaign's scheduleCron field.
cron.schedule('* * * * *',()=>{const state=getState();state.campaigns.filter(c=>c.status==='active'&&c.scheduleCron).forEach(c=>{if(cron.validate(c.scheduleCron)&&cron.schedule(c.scheduleCron,()=>console.log(`Campaign ${c.name} due`),{scheduled:false})) console.log(`Campaign ${c.name} schedule validated`);});});
app.get('/{*splat}',(_,res)=>res.sendFile(path.join(root,'index.html')));
app.listen(port,()=>console.log(`VARDAN.ai running at http://localhost:${port}`));
