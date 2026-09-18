'use strict';
(async function openFieldViewer(){
 const root=new URL('./',document.currentScript.src);
 const query=new URLSearchParams(location.search);
 const touchTablet=/iPad|Android|iPhone|iPod/i.test(navigator.userAgent)||(navigator.platform==='MacIntel'&&navigator.maxTouchPoints>1)||matchMedia('(pointer: coarse)').matches;
 const compact=query.get('view')==='desktop'?false:document.documentElement.dataset.viewer==='light'||touchTablet||query.get('view')==='mobile';
 const source=new URL(compact?'../iseongsanseong-live/light/index.html':'../iseongsanseong-live/index.html',root);
 const map=new URL('../iseongsanseong-live/base.jpg',root);
 const revision=new URL('./revision-v28.js?v=28.0',root);
 const correction=new URL('./roster-v28-1.js?v=28.1',root);
 const roster=new URL('./roster-sync-v28.js?v=28.1',root);
 const title='2026 하남이성산성문화제 | 현장 배치도';
 const controller=new AbortController();const timer=setTimeout(()=>controller.abort(),30000);
 try{
  const fetchText=async url=>{const r=await fetch(url,{cache:'no-store',credentials:'same-origin',signal:controller.signal});if(!r.ok||new URL(r.url).origin!==location.origin)throw new Error('배치도 읽기 오류 '+r.status);return r.text();};
  let [html,patch,duty,sync]=await Promise.all([fetchText(source),fetchText(revision),fetchText(correction),fetchText(roster)]);clearTimeout(timer);
  if(!/<html[\s>]/i.test(html)||!/<script[\s>]/i.test(html)||!patch.includes('fieldRevision28')||!duty.includes('synchronizeFieldRoster')||!sync.includes('syncFieldRoster281'))throw new Error('배치도 형식 오류');
  html=html.replace(/(["'])\/base\.jpg\1/g,(_m,q)=>q+map.pathname+q);
  html=html.replace(/<title>[\s\S]*?<\/title>/i,'<title>'+title+'</title>');
  html=html.replace(/<html([^>]*)>/i,'<html$1 data-field-compact="'+compact+'">');
  const base=new URL('./',source).href.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
  const meta='<base href="'+base+'"><meta name="robots" content="noindex,nofollow,noarchive"><meta name="referrer" content="same-origin"><meta name="application-name" content="하남이성산성문화제 현장 배치도"><meta property="og:title" content="'+title+'"><meta property="og:description" content="시설·운영인력 배치 및 날짜별 임무 확인"><link rel="icon" href="data:image/svg+xml,%3Csvg xmlns=%27http://www.w3.org/2000/svg%27 viewBox=%270 0 64 64%27%3E%3Crect width=%2764%27 height=%2764%27 rx=%2712%27 fill=%27%23253c4f%27/%3E%3Cpath d=%27M14 48V27l18-14 18 14v21M24 48V32h16v16%27 fill=%27none%27 stroke=%27white%27 stroke-width=%275%27/%3E%3C/svg%3E">';
  html=html.replace(/<head([^>]*)>/i,'<head$1>'+meta);
  html=html.replace(/<\/body>/i,'<script>'+(patch+'\n'+duty+'\n'+sync).replace(/<\/script/gi,'<\\/script')+'</script></body>');
  document.open();document.write(html);document.close();
 }catch(error){clearTimeout(timer);const s=document.getElementById('load-status');if(s)s.textContent='배치도를 불러오지 못했습니다. 통신 상태를 확인한 뒤 새로고침해 주세요.';const r=document.getElementById('retry');if(r)r.hidden=false;console.error('배치도 불러오기 실패',error);}
})();
