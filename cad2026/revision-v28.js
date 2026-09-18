(function fieldRevision28(){
'use strict';
if(window.__hcfV28)return;
const compact=document.documentElement.dataset.fieldCompact==='true'||!!document.querySelector('.field-tools');
const clone=x=>JSON.parse(JSON.stringify(x));
const mean=p=>p.reduce((a,b)=>[a[0]+b[0]/p.length,a[1]+b[1]/p.length],[0,0]);
const rect=(x,y,w,d)=>[[x,y],[x+w,y],[x+w,y+d],[x,y+d]];
const shift=(p,x,y)=>p.map(q=>[q[0]+x,q[1]+y,...q.slice(2)]);
const object=id=>scene.objects.find(o=>o.id===id);
const expectedMapSize=clone(scene.background.size);
const calibrationSnapshot=JSON.stringify(scene.background.pixel_to_world);
const originalAward=scene.objects.filter(o=>o.kind==='temporary_chair').map(clone);
const originalCam=clone(object('CAM'));
const allRegistry=()=>Object.values(scene.staff_registry).flat();
const person=id=>allRegistry().find(p=>p.id===id);
const flow='통로에 멈추거나 모여서 관람하지 않도록 이동 방향을 안내하고 계속 지나가도록 유도·혼잡 시 현장 총괄에게 보고';
const setPost=(p,u,v,location,role)=>{if(!p)return;Object.assign(p,{u,v,remote:false,location_label:location,role});delete p.external_area;};
scene.version='v28.0-field';scene.title='2026 하남이성산성문화제 현장 배치도';
scene.field_revision_v28={source:'2026-09-18 사용자 현장 설명',surveyed:false,stage_mat_gap_m:4,ab_gap_model_m:2.5,white_tower:{model_width:6,model_depth:8.4,model_height:7.5,levels:3,seat_layout:'좌우 각 5석 × 8줄의 표현안',dimensions_confirmed:false},registration:'기존 기준점·좌표변환 유지. 측량 정합 확정 아님.'};
// Keep all existing IDs, stage orientation and the reference-image registration.
for(const o of scene.objects){
 if(o.kind==='mat')o.poly=shift(o.poly,0,-1.2);
 if(o.kind==='chair'&&['B1','B3'].includes(o.zone_id||o.block))o.poly=shift(o.poly,0,-1.8);
 if(/^PREP-[1-4]$/.test(o.id)){o.kind='tent';o.peak=3.5;o.color='#e8e5dd';o.name='무대 뒤 몽골텐트 '+o.id.slice(-1);o.banner_text=o.name;}
 if(['O05','O06'].includes(o.id)){o.name='프레스';o.banner_text='프레스';}
 if(/^EXT-AG-(14|15|16)$/.test(o.id)){o.name='팔찌만들기 체험부스';o.banner_text=o.name;}
}
scene.mat_groups.forEach(g=>g.poly=shift(g.poly,0,-1.2));
scene.static_fences_v12.forEach(f=>{if(/^F-MAT-/.test(f.id))f.points=shift(f.points,0,-1.2);if(/^F-LINK-[LR]-B$/.test(f.id))f.points[f.points.length-1][1]-=1.2;});
const rearNames=['NH농협','의료','통합안내소','통합안내소','안전부스'];
rearNames.forEach((name,i)=>{const o=object('REAR-B3-'+(i+1));if(o){o.name='후방 천막 '+(i+1)+' · '+name;o.banner_text=name;}});
scene.objects=scene.objects.filter(o=>!/^EXT-AG-(24|25|26|27)$/.test(o.id));
// White tower occupies the middle of B2; the numerical dimensions below are a model, not an as-built measurement.
scene.objects=scene.objects.filter(o=>!(o.kind==='chair'&&(o.zone_id||o.block)==='B2'));
for(const side of [-1,1])for(let row=0;row<8;row++)for(let col=0;col<5;col++){
 const x=side<0?-6.485+col*.65:3.265+col*.65,y=30.05+row*.95;
 scene.objects.push({id:'V28-B2-'+(side<0?'L':'R')+'-'+row+'-'+col,kind:'chair',name:'관람석',poly:rect(x,y,.62,.65),width:.62,depth:.65,height:.86,color:'#a3a9ae',row:row+1,block:'B2',zone_id:'B2',vip:false,vvip:false,vvip:false,yaw_deg:0});
}
const cam=object('CAM');Object.assign(cam,{kind:'white_tower',name:'화이트타워',poly:rect(-3,29.15,6,8.4),width:6,depth:8.4,height:7.5,note:'사용자 현장 설명: B2 안쪽 진입, 약 3층 구조 및 상부 조명. 표현 치수는 실측 전 가정.'});
for(const o of scene.objects.filter(o=>o.kind==='broadcast_camera'))o.poly=shift(o.poly,0,2.2);
for(const z of scene.seat_zones){
 const chairs=scene.objects.filter(o=>o.kind==='chair'&&(o.zone_id||o.block)===z.id);
 if(!chairs.length)continue;const pts=chairs.flatMap(o=>o.poly),xs=pts.map(p=>p[0]),ys=pts.map(p=>p[1]);
 const x=Math.min(...xs),y=Math.min(...ys),w=Math.max(...xs)-x,d=Math.max(...ys)-y;
 Object.assign(z,{count:chairs.length,poly:rect(x,y,w,d),center:[x+w/2,y+d/2]});
}
scene.counts.chairs=scene.objects.filter(o=>o.kind==='chair').length;
scene.counts.temporary_chairs=originalAward.length;
scene.counts.preparation_canopies=0;scene.counts.preparation_mongol_tents=4;
scene.installation_seating.rendered_chair_count=scene.counts.chairs+originalAward.length;
const commandStaff=scene.staff_registry.city.filter(p=>/합동부스/.test(p.location_label||''));
commandStaff.forEach((p,i)=>setPost(p,-2.45+(i%4)*.63,46.35+Math.floor(i/4)*.63,'후방 천막 5 · 안전부스 내부',(p.role||'상황 접수 및 현장 지원').replaceAll('합동부스','안전부스')));
const f7=person('F07'),f15=person('F15');if(f7?.sunday_post&&f15?.sunday_post){const temp=clone(f7.sunday_post);f7.sunday_post=clone(f15.sunday_post);f15.sunday_post=temp;}
for(const id of ['F07','F12']){const p=person(id);if(p?.sunday_post)p.sunday_post.role='F12·F07 2인 구간 분담 · '+flow;}
const a14=person('A14');setPost(a14,25.35,31.95,'10번 천막–B3 연결 휀스 동측 끝 · 34번과 양 끝 분담',flow);
setPost(person('S34'),20.9,32.2,'10번 천막–B3 연결 휀스 서측 끝 · A14와 양 끝 분담',flow);
setPost(person('G04'),11.14,28.8,'A3·B3 사이 통로',flow+'·무단 좌석 진입 및 돌발 접근 통제');
const air=scene.airbounce.model_poly||scene.airbounce.poly;
for(let i=0;i<5;i++){const t=(i+.5)/5,a=air[1],b=air[2];const p=person('A'+String(i+7).padStart(2,'0'));setPost(p,a[0]+(b[0]-a[0])*t,a[1]+(b[1]-a[1])*t+1.8,'워터슬라이드 옆 · '+['입장 대기','입장부','이용구역 관찰','퇴장부','보호자 통로'][i],p?.role||flow);if(p)p.external_area='airbounce';}
scene.airbounce.label='워터슬라이드';
scene.parking.roadview_detail.building_name='공중화장실·공원관리실';
const canopy={id:'CIVIC-CANOPY',kind:'prep_canopy',name:'자치행정과 캐노피',poly:rect(17.2,7.7,3,3),width:3,depth:3,height:2.15,peak:2.75,color:'#f5f4ee',banner_text:'자치행정과',banner_color:'#6f7c83',banner_height:.25,open_sides:4,inside_table:false,note:'9.19 17:00~18:00 시민의날 사용 / 18:00 브레이크타임 철거. 크기는 표현 가정.'};
const oldRoster=applyStaffRoster;
applyStaffRoster=function(){oldRoster();const sunday=currentStaffDate()==='2026-09-20';
 for(const phase of ['concert','ceremony'])for(const p of scene.staff_phases[phase]){
  if(p.id==='F05'){const q=scene.staff_phases[phase].find(x=>x.id==='A01');if(q)setPost(p,q.u-1.8,q.v,'A01 옆 · 현장 총괄','현장 혼잡·통행 장애 보고 접수·인근 요원 지원 지시·안전부스 및 전문경호와 상황 공유');}
  if(!sunday&&phase==='ceremony'){
   if(p.id==='G03')setPost(p,9.7,7,'03번 기존 위치 · 무대 우측 접근부','시민의날 무대 접근 통제·수상자 이동 시 진행팀 확인 후 통행 지원·출연진 입출차 시 G06과 동행 경호');
   if(p.id==='G06')setPost(p,-9.6,7,'F05 기존 위치 · 무대 좌측 접근부','시민의날 무대 접근 통제·진행팀 이동 지원·출연진 입출차 시 G03과 동행 경호');
   if(p.id==='S03'){const entry=scene.operation_roster.city_hall.people.find(x=>x.person_id===p.person_id);const q=scene.staff_phases[phase].find(x=>x.id==='S09');if(entry?.escort&&q)setPost(p,q.u+1.2,q.v,'09번 옆 · 내빈·수상 지원',entry.assignment);else setPost(p,23.1,32.6,'34번·A14 연결 휀스 통로',flow);}
  }
  if(!sunday&&['S20','S13'].includes(p.id)){const ceremony=phase==='ceremony';setPost(p,ceremony?21.5:20.6,p.id==='S20'?(ceremony?13:21):(ceremony?21:28),'수상자석 옆 · '+(ceremony?'시민의날 진행':'18시 이후 A3 옆'),'수상자석 주변 통로 확보·무단 진입 방지·수상자석 이동 시 함께 재배치');}
  const command=commandStaff.find(r=>r.id===p.id);if(command)setPost(p,command.u,command.v,command.location_label,command.role);
  if(['F01','F02','S16','F20','F21'].includes(p.id)&&p.v>28.4&&p.v<32)p.v=28.8;
 }
 for(const phase of ['concert','ceremony'])scene.staff_phases[phase].sort((a,b)=>Number(!!a.command_only)-Number(!!b.command_only));
};
function applyPhaseObjects(s){
 const ceremony=currentStaffDate()==='2026-09-19'&&$('phase').value==='ceremony';
 s.objects=s.objects.filter(o=>o.kind!=='temporary_chair'&&o.id!=='CIVIC-CANOPY');
 if(currentStaffDate()==='2026-09-19'){
  const dx=ceremony?-.8:-1.7,dy=ceremony?2.5:9.35;
  s.objects.push(...originalAward.map(o=>({...clone(o),poly:shift(o.poly,dx,dy)})));
  const pts=s.objects.filter(o=>o.kind==='temporary_chair').flatMap(o=>o.poly);const lo=[Math.min(...pts.map(p=>p[0])),Math.min(...pts.map(p=>p[1]))],hi=[Math.max(...pts.map(p=>p[0])),Math.max(...pts.map(p=>p[1]))];
  s.awardee_group.center=[(lo[0]+hi[0])/2,(lo[1]+hi[1])/2];s.awardee_group.bounds=[lo,hi];s.temporary_review.existing_bounds={u:[lo[0],hi[0]],v:[lo[1],hi[1]]};
  if(ceremony)s.objects.push(clone(canopy));
 }
}
function addTower(w){
 const parts=w.parts;
 function add(p,c,line=false,width=1){const center=p.reduce((a,b)=>a.map((v,i)=>v+(b[i]||0)/p.length),[0,0,0]);parts.push({p,c,layer:'equipment',line,width,center,radius:Math.max(...p.map(q=>Math.hypot(q[0]-center[0],q[1]-center[1],q[2]-center[2]))),detail:false});}
 function box(x,y,ww,dd,z,h){const a=rect(x,y,ww,dd).map(p=>[...p,z]),b=rect(x,y,ww,dd).map(p=>[...p,z+h]);add(b,'#e5e8e6');for(let i=0;i<4;i++)add([a[i],a[(i+1)%4],b[(i+1)%4],b[i]],i%2?'#aebfc5':'#cdd6d6');}
 for(const x of [-2.9,2.75])for(const y of [29.25,33.3,37.3])box(x,y,.15,.15,0,7.65);
 for(const z of [2.5,5,7.5]){box(-3,29.15,6,8.4,z-.16,.16);for(const x of [-2.9,2.9])for(const y of [29.25,33.3]){add([[x,y,z-2.4],[x,y+4,z-.2]],'#cad7db',true,1.15);add([[x,y+4,z-2.4],[x,y,z-.2]],'#92a8b1',true,1.15);}}
 for(const x of [-2.5,2.1])for(const y of [29.5,36.7]){box(x,y,.4,.4,7.55,.4);add([[x,y,7.75],[x+.4,y,7.75],[x+.4,y,8.05],[x,y,8.05]],'#fff1bf');}
 for(const y of [29.25,37.4])add([[-2.9,y,8.25],[2.9,y,8.25]],'#dfe6e6',true,1.6);
}
const oldCreate=createWorld;
createWorld=function(s){applyPhaseObjects(s);const q={...s,objects:s.objects.filter(o=>o.id!=='CAM')};const w=oldCreate(q);w.grounds=w.grounds.filter(g=>g.layer!=='paths');addTower(w);return w;};
// A single image-space texture keeps both the satellite image and watermark tied to the same reference points.
let mapTexture=null;
function buildMapTexture(){if(mapTexture||!image.naturalWidth)return;const t=document.createElement('canvas');t.width=expectedMapSize[0];t.height=expectedMapSize[1];const c=t.getContext('2d');c.drawImage(image,0,0,t.width,t.height);
 // Clear only the flat grey screenshot border; do not crop/stretch the registered image.
 let border=t.height;try{const pix=c.getImageData(0,Math.max(0,t.height-64),t.width,64).data;for(let y=63;y>=0;y--){let ok=0,n=0;for(let x=48;x<t.width-48;x+=64){const k=(y*t.width+x)*4;const r=pix[k],g=pix[k+1],b=pix[k+2];n++;if(Math.max(r,g,b)-Math.min(r,g,b)<6&&r>115&&r<140)ok++;}if(ok/n>.94)border=t.height-64+y;else break;}if(border<t.height)c.clearRect(0,border,t.width,t.height-border);}catch{}
 c.save();c.font='600 21px "Apple SD Gothic Neo","Malgun Gothic",sans-serif';c.textAlign='center';c.textBaseline='middle';c.lineWidth=2;c.strokeStyle='rgba(12,25,39,.19)';c.fillStyle='rgba(255,255,255,.30)';for(let row=0,y=70;y<t.height;y+=240,row++)for(let x=100+(row%2)*170;x<t.width;x+=340){c.save();c.translate(x,y);c.rotate(-Math.PI/12);c.strokeText('하남문화재단 조형석',0,0);c.fillText('하남문화재단 조형석',0,0);c.restore();}c.restore();mapTexture=t;}
const nativeDraw=ctx.drawImage.bind(ctx);
ctx.drawImage=function(src,...a){if(src===image){buildMapTexture();return nativeDraw(mapTexture||image,0,0,expectedMapSize[0],expectedMapSize[1]);}return nativeDraw(src,...a);};
image.addEventListener('load',()=>{mapTexture=null;buildMapTexture();requestRender();});buildMapTexture();
// Mobile and tablet labels keep facilities and staff; measurement text is desktop-only.
drawLabels=function(layers){if(!$('labels').checked)return;const ext=!!scene.external_areas?.[focus];if(ext||['parking','context'].includes(focus))return;const label=(p,t)=>boxLabel(p,t,'#244c62',compact?10:12);
 if(layers.stage){label([0,.5,2.5],compact?'메인무대':'메인무대 16.3×9.1m');if(!compact){label([-11.65,3.5,10.8],'좌측 사이드타워 W6×H10m');label([11.65,3.5,10.8],'우측 사이드타워 W6×H10m');}}
 if(layers.mats&&(!compact||zoom>1.25))label([0,13,.3],compact?'돗자리존':'돗자리존 · 무대 앞끝 이격 4m');
 if(layers.tents){for(const o of scene.objects){if(o.booth_number&&(!compact||[1,10].includes(o.booth_number)))label([...mean(o.poly),3.8],o.booth_number+'번');if((!compact||zoom>1.7)&&(/REAR-B3/.test(o.id)||['O05','O06','CIVIC-CANOPY'].includes(o.id)))label([...mean(o.poly),3.8],o.banner_text);}}
 if(!compact||zoom>1.4){label([0,33.2,8.5],'화이트타워');if(awardeeSeatsEnabled())label(awardeeAnchor(1.3),'수상자석');if(layers.tents&&(!compact||focus==='dome'))label([...mean(scene.objects.filter(o=>/^PREP-/.test(o.id)).flatMap(o=>o.poly)),4.3],'무대 뒤 몽골텐트 4동');}
};
const oldExternalLabels=drawExternalLabels;
drawExternalLabels=function(){if(!$('labels').checked)return;if(focus==='airbounce')boxLabel([...scene.airbounce.label_uv,2.8],'워터슬라이드','#176579',12);else oldExternalLabels();const b=scene.parking.roadview_detail.building;if(b)boxLabel([...mean(b),4.3],'공중화장실·공원관리실','#244c62',compact?10:12);};
if(compact){drawMeasurements=function(){};drawAisleDims=function(){};}
 drawDrawingStamp=function(){ctx.save();ctx.font='10px "Apple SD Gothic Neo","Malgun Gothic",sans-serif';ctx.textAlign='left';ctx.fillStyle='#5c7485';ctx.fillText('하남문화재단 조형석 · v28',12,trans.h-56);ctx.restore();};
const stamp=document.querySelector('.title span');if(stamp)stamp.textContent='현장 배치 · v28';
const badge=document.getElementById('layoutBadge');if(badge)badge.textContent='현장 배치 · v28';
if($('phase').options[0])$('phase').options[0].textContent='18시 이후·공연';
if($('phase').options[1])$('phase').options[1].textContent='시민의날 17~18시';
// Pointer handlers: single-finger orbit; two-finger pinch and pan. View is never reset by touch or date/phase changes.
const mouseDown=canvas.onpointerdown,mouseMove=canvas.onpointermove,mouseUp=canvas.onpointerup;
let gesture=null,hadPinch=false;
const local=e=>{const r=canvas.getBoundingClientRect();return{x:e.clientX-r.left,y:e.clientY-r.top};};
const pair=()=>{const a=[...touchPoints.values()];return{x:(a[0].x+a[1].x)/2,y:(a[0].y+a[1].y)/2,d:Math.max(1,Math.hypot(a[0].x-a[1].x,a[0].y-a[1].y))};};
canvas.style.touchAction='none';canvas.style.userSelect='none';canvas.style.webkitUserSelect='none';
canvas.onpointerdown=function(e){if(e.pointerType!=='touch')return mouseDown?.(e);e.preventDefault();touchPoints.set(e.pointerId,local(e));try{canvas.setPointerCapture(e.pointerId);}catch{}if(touchPoints.size===1){hadPinch=false;gesture={kind:'orbit',...local(e),az,el};}else if(touchPoints.size===2){hadPinch=true;gesture={kind:'pinch',...pair(),zoom,panX,panY};}drag=null;};
canvas.onpointermove=function(e){if(e.pointerType!=='touch')return mouseMove?.(e);if(!touchPoints.has(e.pointerId)||!gesture)return;e.preventDefault();touchPoints.set(e.pointerId,local(e));if(touchPoints.size>=2&&gesture.kind==='pinch'){const p=pair(),ratio=p.d/gesture.d;zoom=Math.max(.6,Math.min(5,gesture.zoom*ratio));const k=zoom/gesture.zoom;panX=gesture.panX+p.x-gesture.x+(1-k)*(gesture.x-canvas.clientWidth/2-gesture.panX);panY=gesture.panY+p.y-gesture.y+(1-k)*(gesture.y-canvas.clientHeight/2-gesture.panY);moving=true;requestRender();return;}const p=local(e),dx=p.x-gesture.x,dy=p.y-gesture.y;if(Math.hypot(dx,dy)<5&&!moving)return;AZ.value=((gesture.az+dx*.4+540)%360)-180;EL.value=Math.max(20,Math.min(90,gesture.el-dy*.25));moving=true;update();};
function finishTouch(e){if(e.pointerType!=='touch')return mouseUp?.(e);touchPoints.delete(e.pointerId);if(touchPoints.size===1){const p=[...touchPoints.values()][0];gesture={kind:'orbit',...p,az,el};}else if(!touchPoints.size){gesture=null;moving=false;drag=null;pinch=null;requestRender();}}
canvas.onpointerup=finishTouch;canvas.onpointercancel=function(e){touchPoints.clear();gesture=null;moving=false;drag=null;pinch=null;requestRender();};
window.addEventListener('blur',()=>{touchPoints.clear();gesture=null;moving=false;drag=null;pinch=null;});
const home=()=>{touchPoints.clear();gesture=null;setView('all',45,35);};
const style=document.createElement('style');style.textContent='#fieldHome{position:fixed;left:10px;bottom:calc(10px + env(safe-area-inset-bottom));z-index:29;min-height:38px;background:#ffffffeb;border:1px solid #a6b7c4;border-radius:5px;color:#284b63;font:12px "Apple SD Gothic Neo","Malgun Gothic",sans-serif;padding:0 12px}#fieldNotice{position:fixed;inset:0;z-index:10000;display:grid;place-items:center;padding:24px;background:#142a3ddb;backdrop-filter:blur(5px);-webkit-backdrop-filter:blur(5px)}#fieldNotice .card{max-width:440px;width:100%;padding:28px 24px;border:1px solid #9badbc;border-radius:12px;background:#f8fafc;color:#223b4e;box-shadow:0 18px 65px #0004;text-align:center;font-family:"Apple SD Gothic Neo","Malgun Gothic",sans-serif}#fieldNotice h2{font-size:21px;margin:12px 0 18px;line-height:1.5}#fieldNotice p{font-size:14px;line-height:1.8;margin:10px 0}#fieldNotice small{font-size:12px;color:#61788a}#fieldNotice .bar{height:3px;background:#355f79;margin:20px 0 0;transform-origin:left;animation:hcfNotice 2.5s linear forwards}@keyframes hcfNotice{to{transform:scaleX(0)}}';document.head.append(style);
if(compact&&!document.getElementById('fieldHome')){const h=document.createElement('button');h.id='fieldHome';h.type='button';h.textContent='처음 각도';h.title='한 손가락 회전 · 두 손가락 확대·이동 · 요원 터치 상세';h.onclick=home;document.body.append(h);}
let notice=document.getElementById('fieldNotice');if(!notice){notice=document.createElement('div');notice.id='fieldNotice';notice.setAttribute('role','dialog');notice.setAttribute('aria-modal','true');notice.setAttribute('aria-labelledby','noticeTitle');notice.innerHTML='<div class="card"><small>2026 하남이성산성문화제 · 현장 배치도</small><h2 id="noticeTitle">무단 복제·배포 금지</h2><p>본 자료는 행사 운영 및 안전관리 목적으로만 사용합니다.<br>무단 복제·수정·재배포 및 타 행사 전용을 금합니다.</p><p><strong>하남문화재단 조형석</strong></p><small>약 2.5초 후 배치도가 열립니다.</small><div class="bar"></div></div>';document.body.append(notice);}const shown=performance.now();setTimeout(()=>{notice.remove();window.__hcfV28.noticeDuration=performance.now()-shown;},2500);
if(compact)home();rebuildOperationalWorld();requestRender();
window.__hcfV28={version:'v28.0',compact,initialView:{az:45,el:35},getState:()=>({az,el,zoom,panX,panY,focus,moving,pointers:touchPoints.size,stats,calibrationUnchanged:JSON.stringify(scene.background.pixel_to_world)===calibrationSnapshot,mapTextureReady:!!mapTexture}),reset:home};
})();
