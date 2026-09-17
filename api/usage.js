const KEY_TOTAL = 'notetong:total-visits';
const KEY_DAY_PREFIX = 'notetong:daily-visitors:';

function credentials() {
  return { url:process.env.UPSTASH_REDIS_REST_URL||process.env.KV_REST_API_URL, token:process.env.UPSTASH_REDIS_REST_TOKEN||process.env.KV_REST_API_TOKEN };
}
function koreaDateKey() {
  const parts=new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Seoul',year:'numeric',month:'2-digit',day:'2-digit'}).formatToParts(new Date());
  const map={};for(const p of parts)map[p.type]=p.value;return map.year+'-'+map.month+'-'+map.day;
}
async function pipeline(commands) {
  const config=credentials();if(!config.url||!config.token)throw new Error('STATS_STORAGE_NOT_CONFIGURED');
  const response=await fetch(config.url.replace(/\/$/,'')+'/pipeline',{method:'POST',headers:{Authorization:'Bearer '+config.token,'Content-Type':'application/json'},body:JSON.stringify(commands)});
  if(!response.ok)throw new Error('UPSTASH_'+response.status);return response.json();
}
function send(res,status,body){res.setHeader('Cache-Control','no-store, max-age=0');res.setHeader('Content-Type','application/json; charset=utf-8');res.status(status).json(body);}

module.exports=async function handler(req,res){
  if(req.method!=='GET'&&req.method!=='POST'){res.setHeader('Allow','GET, POST');return send(res,405,{error:'METHOD_NOT_ALLOWED'});}
  try{
    const dayKey=KEY_DAY_PREFIX+koreaDateKey();let commands;
    if(req.method==='POST'){
      const token=req.body&&String(req.body.visitorToken||'');
      if(!/^[a-f0-9]{64}$/.test(token))return send(res,400,{error:'INVALID_VISITOR_TOKEN'});
      commands=[['SADD',dayKey,token],['EXPIRE',dayKey,259200],['INCR',KEY_TOTAL],['SCARD',dayKey]];
      const result=await pipeline(commands);
      return send(res,200,{todayVisitors:Number(result[3]&&result[3].result)||0,totalVisits:Number(result[2]&&result[2].result)||0});
    }
    commands=[['SCARD',dayKey],['GET',KEY_TOTAL]];
    const result=await pipeline(commands);
    return send(res,200,{todayVisitors:Number(result[0]&&result[0].result)||0,totalVisits:Number(result[1]&&result[1].result)||0});
  }catch(error){const missing=error&&error.message==='STATS_STORAGE_NOT_CONFIGURED';return send(res,missing?503:500,{error:missing?'STATS_STORAGE_NOT_CONFIGURED':'STATS_UNAVAILABLE'});}
};
