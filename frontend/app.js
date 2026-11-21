'use strict';
const $ = id => document.getElementById(id);
let pendingSpeech = null, selectedAudio = null;
let current = null, running = false, lastHealth = null, stepNodes = new Map(), runEvents = [];
const conversations = new Map();
const statusLabels = {running:'در حال اجرا',success:'تکمیل شد',failed:'ناموفق'};
const toolLabels = {sales_breakdown:'تفکیک فروش بر اساس کانال',sales_order_count:'شمارش سفارش‌ها',sales_revenue:'محاسبه درآمد فروش',business_semantics:'خواندن تعریف شاخص‌ها'};
function node(tag,text,parent,cls){const n=document.createElement(tag);if(text!==null)n.textContent=text;if(cls)n.className=cls;if(parent)parent.append(n);return n;}
function jsonBlock(parent,label,value){const d=node('details',null,parent);node('summary',label,d);node('pre',JSON.stringify(unpack(value),null,2),d);return d;}
function setStatus(text,error=false){$('status').textContent=text;$('status').classList.toggle('error',error);}
function time(){return new Date().toLocaleTimeString('fa-IR',{hour:'2-digit',minute:'2-digit'});}
async function api(path,options={}){options.headers={...options.headers};if($('token').value)options.headers.Authorization='Bearer '+$('token').value;const response=await fetch(path,options);if(!response.ok){let data;try{data=await response.json();}catch{data={};}throw Error(data.message||'ارتباط با سرویس کامل نشد ('+response.status+')');}return response;}
function scrollBottom(){const scroller=$('chat-scroll');scroller.scrollTo({top:scroller.scrollHeight,behavior:'smooth'});}
function message(role,text){$('welcome').hidden=true;const article=node('article',null,$('messages'),'message '+role),head=node('div',null,article,'message-header');node('span',role==='user'?'ش':'◈',head,'message-avatar');node('strong',role==='user'?'شما':role==='transcript'?'متن صوت · Whisper':'MobinAid',head);node('span',time(),head,'message-time');const body=node('div',null,article,'message-body');renderAnswer(body,text);scrollBottom();return article;}
function unpack(value){if(value&&typeof value==='object'&&typeof value.text==='string'){try{return JSON.parse(value.text);}catch{}}return value;}
function renderAnswer(parent,text){
 // Mail metadata is rendered once as actionable cards, not repeated in prose.
 if(text.startsWith('بازهٔ بررسی')&&text.includes('ایمیل دریافتی در این صفحه پیدا شد:')){
  const start=text.indexOf('\n\n1. '),end=text.indexOf('\n\nاین فهرست موضوع',start);
  if(start>=0&&end>start)text=text.slice(0,start)+text.slice(end).replace('این فهرست موضوع و مشخصات ایمیل‌هاست؛ بدنهٔ ایمیل‌ها در این مرحله خوانده نشده است.','مشخصات ایمیل‌ها در کارت‌های زیر است؛ برای خواندن متن روی «متن کامل» بزن.');
 }
 const lines=text.split('\n');let pending=[];
 const flush=()=>{if(pending.length){node('div',pending.join('\n'),parent,'answer-prose');pending=[];}};
 for(let i=0;i<lines.length;i++){
  if(lines[i].trim().startsWith('|')&&i+1<lines.length&&/^\s*\|[\s:|\-]+\|\s*$/.test(lines[i+1])){
   flush();const wrapper=node('div',null,parent,'table-wrap'),table=node('table',null,wrapper,'answer-table');
   const cells=line=>line.trim().replace(/^\||\|$/g,'').split('|').map(x=>x.trim());
   const head=node('thead',null,table),tr=node('tr',null,head);cells(lines[i]).forEach(c=>node('th',c,tr));
   const body=node('tbody',null,table);i+=2;
   while(i<lines.length&&lines[i].trim().startsWith('|')){const row=node('tr',null,body);cells(lines[i]).forEach(c=>node('td',c,row));i++;}i--;
  }else pending.push(lines[i]);
 }flush();
}
function evidence(event,parent){
 const d=node('details',null,parent,'inline-evidence');node('summary','منبع واقعی · '+event.tool+' · '+event.duration_ms+' ms',d);
 const value=unpack(event.result),rows=Array.isArray(value?.data)?value.data:Array.isArray(value?.data?.customers)?value.data.customers:null;
 if(rows?.length&&rows.length<=200&&rows[0]&&typeof rows[0]==='object'){
  const keys=Object.keys(rows[0]).filter(k=>rows.every(r=>r&&typeof r[k]!=='object')).slice(0,8);
  const wrapper=node('div',null,d,'table-wrap'),table=node('table',null,wrapper),head=node('tr',null,table);
  const labels={customer_id:'شناسه مشتری',name:'نام',city:'شهر',paid_orders:'سفارش موفق',lifetime_value:'مجموع خرید',orders:'سفارش',revenue:'مبلغ',category:'دسته',count:'تعداد',group:'گروه'};
  keys.forEach(k=>node('th',labels[k]||k,head));
  for(const row of rows){const tr=node('tr',null,table);keys.forEach(k=>node('td',typeof row[k]==='number'?row[k].toLocaleString('fa-IR'):String(row[k]??''),tr));}
 }
 jsonBlock(d,'جزئیات فنی · JSON خوانا',value);
}
function renderEvidence(events,parent){
 const mailRows=new Map();
 const reads=events.filter(e=>e.tool==='outlook_read_email'&&e.status==='success'&&e.show_cards!==false);const searches=events.filter(e=>['outlook_search_emails','outlook_selected_messages'].includes(e.tool)&&e.status==='success'&&e.show_cards!==false);const cardEvents=reads.length?reads:searches.slice(-1);
 for(const event of cardEvents){const raw=unpack(event.result),data=raw?.data||raw;if(['outlook_search_emails','outlook_selected_messages'].includes(event.tool))for(const row of data?.items||[])mailRows.set(row.id,row);if(event.tool==='outlook_read_email'&&data?.id)mailRows.set(data.id,data);}
 if(mailRows.size){const cards=node('section',null,parent,'mail-cards');const conversation=current;for(const row of mailRows.values()){
  const card=node('article',null,cards,'mail-card');node('strong',row.subject||'بدون موضوع',card);node('div',row.sender||'',card,'mail-meta');
  const when=row.sent_utc||row.received_utc;if(when){const date=new Date(when);if(!Number.isNaN(date.getTime()))node('small',new Intl.DateTimeFormat('fa-IR',{timeZone:'Asia/Tehran',dateStyle:'medium',timeStyle:'short'}).format(date),card);}
  const actions=node('div',null,card,'mail-actions');for(const label of ['متن کامل','پیوست‌ها','خلاصه']){const button=node('button',label,actions);button.type='button';button.onclick=()=>{if(running)return;if(current!==conversation){setStatus('این کارت متعلق به گفت‌وگوی دیگری است؛ همان گفت‌وگو را باز کن.',true);return;}$('text').value='ایمیل '+row.id+' '+label;$('form').requestSubmit();};}
 }}
 const calls=events.filter(e=>e.tool!=='outlook_selected_messages'&&!e.cached);
 if(calls.length>1){const group=node('details',null,parent,'evidence-group');node('summary',calls.length.toLocaleString('fa-IR')+' فراخوانی واقعی · مشاهدهٔ منابع',group);for(const e of calls)evidence(e,group);}
 else for(const e of calls)evidence(e,parent);
}
function toolsSummary(mcp){return (mcp?.tools?.length||0).toLocaleString('fa-IR')+' ابزار مجاز از '+(mcp?.discovered_total??'—').toLocaleString('fa-IR')+' ابزار کشف‌شده';}
function connection(name,subtitle,endpoint,available,icon,messageText){const card=node('div',null,$('connections'),'connection'),head=node('div',null,card,'connection-head');node('span',icon,head,'connection-icon');const title=node('div',null,head);node('div',name,title,'connection-name');node('div',subtitle,title,'connection-subtitle');const badge=node('span',null,head,'connection-state '+(available?'good':'bad'));node('span','',badge,'dot '+(available?'green':'amber'));node('span',available?'آماده':'نیاز به بررسی',badge);node('div',endpoint||'آدرس دریافت نشده',card,'endpoint');if(messageText)node('p',messageText,card,'connection-message');}
async function health(){if(running)return;$('refresh').disabled=true;$('checked-at').textContent='در حال بررسی…';try{const data=await(await api('/api/health')).json();lastHealth=data;$('model-name').textContent=data.model;$('model-dot').className='dot '+(data.ollama?.available?'green':'amber');$('health').textContent=JSON.stringify(data,null,2);$('connections').replaceChildren();connection('Ollama',data.model,data.llm_url,data.ollama?.available,'◈',data.ollama?.message);connection('Company Micro APIs', 'MCP · '+toolsSummary(data.mcp),data.mcp_url,data.mcp?.available,'⌘',data.mcp?.message);connection('Persian Whisper', 'تبدیل صوت فارسی · بارگذاری هنگام نیاز', 'stdio · transcribe_audio', data.speech?.available, '♫',data.speech?.message);const tools=data.mcp?.tools||[];$('tool-count').textContent=tools.length.toLocaleString('fa-IR');$('available-count').textContent=tools.length.toLocaleString('fa-IR');$('tools-list').replaceChildren();for(const definition of data.mcp?.definitions||[]){const f=definition.function,card=node('div',null,$('tools-list'),'tool-definition');node('bdi',f.name,card);node('p',toolLabels[f.name]||f.description,card);jsonBlock(card,'قرارداد ورودی ابزار',f.parameters);}if(!data.mcp?.available)node('p','ابزار تأییدشده‌ای در دسترس نیست.',$('tools-list'));$('checked-at').textContent='بررسی '+time();if(!runEvents.length&&!selectedAudio){if(!data.ollama?.available)setStatus(data.ollama?.message||'مدل آماده نیست.',true);else if(!data.mcp?.available)setStatus('اتصال MCP شرکت آماده نیست.',true);else setStatus('مدل و ابزارهای شرکت آماده‌اند');}}catch(error){$('checked-at').textContent='ارتباط ناموفق';$('model-dot').className='dot red';$('connections').replaceChildren();node('p',error.message,$('connections'),'connection-message');setStatus(error.message,true);}finally{$('refresh').disabled=false;}}
function resetTrace(){stepNodes=new Map();runEvents=[];$('trace').replaceChildren();$('run-state').textContent='در حال اجرا';}
function stageTitle(event){if(event.plan||event.identity||event.route==='outlook')return 'برنامه‌ریزی و تطبیق Outlook';if(event.stage==='routing')return event.route==='company'?'انتخاب ابزار مرتبط با درخواست':'گفت‌وگو · بدون اجرای ابزار شرکت';if(event.stage==='speech')return 'تبدیل صوت به متن فارسی';if(event.stage==='report')return 'محاسبهٔ قطعی گزارش از دادهٔ ابزار';if(event.stage==='provider')return 'بررسی مدل محلی';if(event.stage==='mcp')return 'اتصال و کشف ابزارهای MCP';if(event.stage==='model')return 'فراخوانی مدل · مرحله '+String(event.iteration||1).toLocaleString('fa-IR');if(event.stage==='tool')return toolLabels[event.tool]||'اجرای ابزار';if(event.stage==='final')return 'پاسخ نهایی آماده شد';return 'اجرا متوقف شد';}
function trace(event,replay=false){if(!replay)runEvents.push(event);const key=event.id||event.stage;let step=stepNodes.get(key);if(!step){step=node('div',null,$('trace'),'trace-step');stepNodes.set(key,step);}step.className='trace-step '+(event.stage==='error'?'failed':event.stage==='final'?'success':event.status||'running');step.replaceChildren();node('h3',stageTitle(event),step);if(event.model)node('bdi',event.model,step);if(event.tool)node('bdi',event.tool,step);const meta=node('div',null,step,'step-meta');node('span',event.stage==='error'?'خطا':event.stage==='final'?'پایان اجرا':statusLabels[event.status]||'',meta);if(event.duration_ms!==undefined)node('bdi',(event.duration_ms/1000).toLocaleString('fa-IR',{maximumFractionDigits:2})+' ثانیه',meta);if(event.endpoint)node('div',event.endpoint,step,'endpoint');if(event.operation)node('div',event.operation,step,'operation');if(event.run_id){step.dataset.runId=event.run_id;step.title='Run '+event.run_id+' · #'+event.sequence+' · '+event.elapsed_ms+' ms';}if(event.tools)node('div',event.tools.length.toLocaleString('fa-IR')+' ابزار مجاز · '+event.tools.join('، '),step,'step-tools');if(event.plan)jsonBlock(step,'برنامهٔ ساختاریافتهٔ درخواست',event.plan);if(event.identity)node('div','هویت تطبیق‌شده: '+event.identity,step);if(event.arguments)jsonBlock(step,'ورودی فراخوانی',event.arguments);if(event.result)jsonBlock(step,'خروجی واقعی ابزار',event.result);if(event.code)node('bdi',event.code,step);if(event.message)node('p',event.message,step,'connection-message');if(event.stage==='error'){for(const n of stepNodes.values())if(n.classList.contains('running')){n.classList.remove('running');n.classList.add('failed');const label=n.querySelector('.step-meta span');if(label)label.textContent='کامل نشد';}$('run-state').textContent='متوقف شد';}if(event.stage==='final')$('run-state').textContent='تکمیل شد'+(event.elapsed_ms!==undefined?' · '+(event.elapsed_ms/1000).toLocaleString('fa-IR',{maximumFractionDigits:1})+' ثانیه':'');}
function selectHistory(){for(const button of $('history').children)button.classList.toggle('selected',button.dataset.id===current);}
async function fresh(){if(running)return false;try{current=(await(await api('/api/conversations',{method:'POST'})).json()).id;conversations.set(current,{title:'گفت‌وگوی تازه',trace:[]});const key=current,button=node('button','گفت‌وگوی تازه',$('history'));button.dataset.id=key;button.onclick=()=>restore(key);$('messages').replaceChildren();$('welcome').hidden=false;$('trace').replaceChildren();node('p','هنوز درخواستی اجرا نشده است.',$('trace'),'trace-empty');$('run-state').textContent='بدون اجرا';selectHistory();setStatus('گفت‌وگوی تازه آماده است');$('text').focus();return true;}catch(error){setStatus(error.message,true);return false;}}
async function restore(key){if(running)return;try{const data=await(await api('/api/conversations/'+key)).json();current=key;$('messages').replaceChildren();$('welcome').hidden=data.messages.length>0;for(const m of data.messages)if(m.role==='user'||(m.role==='assistant'&&!m.tool_calls))message(m.role,m.content);if(data.evidence?.length){const group=node('div',null,$('messages'));renderEvidence(data.evidence,group);}resetTrace();const stored=conversations.get(key)?.trace||[];for(const e of stored)trace(e,true);runEvents=stored.slice();$('run-state').textContent=stored.length?'اجرای قبلی':'بدون اجرا';selectHistory();setStatus('گفت‌وگو بازیابی شد');}catch(error){setStatus(error.message,true);}}
function inspect(force=false){if(matchMedia('(max-width:1150px)').matches){$('inspector').classList.toggle('mobile-open',force||!$('inspector').classList.contains('mobile-open'));}else if(force)document.body.classList.remove('inspect-hidden');else document.body.classList.toggle('inspect-hidden');}
async function submit(event){event.preventDefault();if(running)return;if(selectedAudio){await uploadAudio(selectedAudio);return;}if(typeof prepareAttachments==='function'&&!await prepareAttachments())return;const draftLink=$('link-url').value.trim();const promptText=$('text').value.trim();const text=(draftLink?((promptText||'این لینک را بررسی کن.')+'\n'+draftLink):promptText)||(typeof pendingFiles!=='undefined'&&pendingFiles.length?'پیوست‌ها را بررسی کن.':'');if(!text)return;if(text.length>4000){setStatus('متن و پرامپت از ظرفیت ۴۰۰۰ نویسه بیشتر است؛ متن را کوتاه‌تر کن. چیزی پنهانی حذف نمی‌شود.',true);return;}if(!current&&!await fresh())return;const item=conversations.get(current);if(item?.title==='گفت‌وگوی تازه'){item.title=text.slice(0,38);const b=Array.from($('history').children).find(b=>b.dataset.id===current);if(b)b.textContent=item.title;}running=true;$('attach-audio').disabled=true;$('send').disabled=true;$('new').disabled=true;message('user',pendingSpeech?.prompt||text);resetTrace();if(pendingSpeech){trace(pendingSpeech);pendingSpeech=null;}setStatus('در حال بررسی مدل و اتصال‌های واقعی…');let ended=false;try{const response=await api('/api/chat',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({conversation_id:current,text,request_key:crypto.randomUUID(),attachment_ids:typeof pendingFiles!=='undefined'?pendingFiles.map(f=>f.id).filter(Boolean):[]})});const reader=response.body.getReader(),decoder=new TextDecoder();let buffer='';while(true){const {value,done}=await reader.read();buffer+=decoder.decode(value||new Uint8Array(),{stream:!done});let index;while((index=buffer.indexOf('\n'))>=0){const line=buffer.slice(0,index);buffer=buffer.slice(index+1);if(!line)continue;const e=JSON.parse(line);trace(e);if(e.stage==='final'){ended=true;const article=message('assistant',e.answer);renderEvidence(e.evidence||[],article);setStatus('پاسخ کامل شد · شواهد ابزار کنار پاسخ در دسترس است');$('text').value='';if(typeof clearFiles==='function')clearFiles();$('link-url').value='';$('link-bar').hidden=true;}else if(e.stage==='error'){ended=true;message('error',e.message+'\nکد خطا: '+e.code);setStatus(e.message,true);}else setStatus(stageTitle(e)+' · '+(statusLabels[e.status]||''));}if(done)break;}if(!ended)throw Error('ارتباط پیش از دریافت نتیجهٔ نهایی قطع شد.');}catch(error){trace({stage:'error',code:'connection_error',message:error.message});message('error',error.message);setStatus(error.message,true);}finally{running=false;$('attach-audio').disabled=false;$('send').disabled=false;$('new').disabled=false;if(item)item.trace=runEvents.slice();scrollBottom();}}
$('form').addEventListener('submit',submit);$('text').addEventListener('keydown',event=>{if(event.key==='Enter'&&!event.shiftKey&&!event.isComposing){event.preventDefault();$('form').requestSubmit();}});$('new').onclick=()=>{if(!running){clearSelectedAudio();if(typeof clearFiles==='function')clearFiles();$('link-url').value='';$('link-bar').hidden=true;fresh();}};$('refresh').onclick=health;$('apply-token').onclick=health;$('inspect-toggle').onclick=()=>inspect();$('connections-nav').onclick=()=>{inspect(true);$('inspector').querySelector('.inspector-body').scrollTop=0;};$('chat-nav').onclick=()=>{$('text').focus();if(matchMedia('(max-width:1150px)').matches)$('inspector').classList.remove('mobile-open');};document.querySelectorAll('[data-prompt]').forEach(button=>button.onclick=()=>{$('text').value=button.dataset.prompt;$('text').focus();});health();setInterval(()=>{if(!document.hidden&&!running)health();},30000);

async function uploadAudio(file){
 if(!file||running)return;
 const extension='.'+file.name.split('.').pop().toLowerCase();
 if(!['.wav','.mp3','.ogg'].includes(extension)){setStatus('فقط WAV، MP3 یا OGG پذیرفته می‌شود.',true);return;}
 if(file.size>25*1024*1024||!file.size){setStatus('فایل باید غیرخالی و حداکثر ۲۵ مگابایت باشد.',true);return;}
 if(!current&&!await fresh())return;
 const draft=$('text').value.trim();running=true;$('send').disabled=true;$('new').disabled=true;$('attach-audio').disabled=true;
 resetTrace();const start=performance.now();trace({stage:'speech',status:'running',model:lastHealth?.speech?.model||'Whisper',operation:'MCP stdio · transcribe_audio'});
 setStatus('در حال تبدیل '+file.name+' به متن فارسی…');
 let transcript=null;
 try{
  const response=await api('/api/audio?conversation_id='+encodeURIComponent(current)+'&extension='+encodeURIComponent(extension),{method:'POST',headers:{'Content-Type':'application/octet-stream'},body:file});
  const result=await response.json();if(!result.text?.trim())throw Error('متن صوت دریافت نشد.');
  const event={stage:'speech',status:'success',prompt:draft,model:result.model,duration_ms:Math.round(performance.now()-start),operation:'MCP stdio · transcribe_audio',result};
  trace(event);pendingSpeech=event;transcript=result.text;message('transcript',result.text);
  setStatus('متن فارسی دریافت شد؛ در حال ارسال به ایجنت…');
 }catch(error){pendingSpeech=null;trace({stage:'error',code:'stt_failed',message:error.message});message('error',error.message);setStatus(error.message,true);}
 finally{running=false;$('send').disabled=false;$('new').disabled=false;$('attach-audio').disabled=false;$('audio-file').value='';}
 if(transcript){clearSelectedAudio();$('text').value=draft?draft+'\n\nمتن فایل صوتی:\n'+transcript:transcript;await submit({preventDefault(){}});}
}
$('attach-audio').onclick=()=>{if(!running)$('audio-file').click();};
function clearSelectedAudio(){selectedAudio=null;$('audio-chip').hidden=true;$('audio-file').value='';$('text').required=true;}
$('remove-audio').onclick=clearSelectedAudio;
$('audio-file').onchange=()=>{const file=$('audio-file').files[0];if(!file)return;
 if(!/\.(wav|mp3|ogg)$/i.test(file.name)||!file.size||file.size>25*1024*1024){setStatus('فایل WAV، MP3 یا OGG غیرخالی تا ۲۵ مگابایت انتخاب کن.',true);return;}
 selectedAudio=file;$('audio-name').textContent=file.name;$('audio-chip').hidden=false;$('text').required=false;
 setStatus('فایل آماده است؛ توضیح دلخواهت را بنویس و ارسال را بزن.');$('text').focus();};


