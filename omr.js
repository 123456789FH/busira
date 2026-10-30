/* Basirah: dependency-free optical mark reader for its own A4 template. */
(function(root,factory){const api=factory();if(typeof module==='object'&&module.exports)module.exports=api;else root.BasirahOMR=api;})(typeof globalThis!=='undefined'?globalThis:this,function(){
  'use strict';
  const WIDTH=800,HEIGHT=1132,MARKERS=[{x:42,y:190},{x:758,y:190},{x:758,y:1050},{x:42,y:1050}];
  function layout(count,options=4){
    if(!Number.isInteger(count)||count<1||count>100||!Number.isInteger(options)||options<2||options>5)throw Error('Invalid sheet dimensions');
    const columns=count<=20?1:count<=50?2:4,rows=Math.ceil(count/columns),cw=664/columns,step=Math.min(35,744/rows);
    return Array.from({length:count},(_,i)=>{const col=Math.floor(i/rows),row=i%rows,right=732-col*cw;const gap=columns===4?23:columns===2?40:57;return {i,x:right,y:263+row*step,r:columns===4?8.1:10,bubbles:Array.from({length:options},(_,j)=>({x:right-32-j*gap,y:263+row*step}))};});
  }
  function solve(matrix,vector){
    const n=vector.length,a=matrix.map((row,i)=>[...row,vector[i]]);
    for(let col=0;col<n;col++){
      let pivot=col;for(let r=col+1;r<n;r++)if(Math.abs(a[r][col])>Math.abs(a[pivot][col]))pivot=r;
      if(Math.abs(a[pivot][col])<1e-10)throw Error('تعذّر ضبط الزوايا. حددي مراكز المربعات الأربعة بدقة.');
      [a[col],a[pivot]]=[a[pivot],a[col]];const d=a[col][col];for(let c=col;c<=n;c++)a[col][c]/=d;
      for(let r=0;r<n;r++){if(r===col)continue;const f=a[r][col];for(let c=col;c<=n;c++)a[r][c]-=f*a[col][c];}
    }return a.map(r=>r[n]);
  }
  function homography(src,dst){
    const a=[],b=[];src.forEach(({x,y},i)=>{const u=dst[i].x,v=dst[i].y;a.push([x,y,1,0,0,0,-u*x,-u*y],[0,0,0,x,y,1,-v*x,-v*y]);b.push(u,v);});return [...solve(a,b),1];
  }
  function project(h,x,y){const d=h[6]*x+h[7]*y+1;return {x:(h[0]*x+h[1]*y+h[2])/d,y:(h[3]*x+h[4]*y+h[5])/d};}
  function gray(image){const a=new Uint8Array(image.width*image.height),d=image.data;for(let i=0;i<a.length;i++)a[i]=Math.round((d[i*4]*.299+d[i*4+1]*.587+d[i*4+2]*.114)*(d[i*4+3]/255)+255*(1-d[i*4+3]/255));return a;}
  function validCorners(p,w,h){
    if(!p||p.length!==4||p.some(a=>!Number.isFinite(a.x)||!Number.isFinite(a.y)||a.x<0||a.y<0||a.x>=w||a.y>=h))return false;
    let area=0;for(let i=0;i<4;i++){const a=p[i],b=p[(i+1)%4],c=p[(i+2)%4];if((b.x-a.x)*(c.y-b.y)-(b.y-a.y)*(c.x-b.x)<=0)return false;area+=a.x*b.y-b.x*a.y;}
    return area/2>w*h*.12;
  }
  function detectMarkers(image){
    const w=image.width,h=image.height,g=gray(image),seen=new Uint8Array(w*h),queue=new Int32Array(w*h),min=Math.max(5,Math.min(w,h)*.008),max=Math.min(w,h)*.065,candidates=[];
    for(let start=0;start<g.length;start++){
      if(seen[start]||g[start]>92)continue;let head=0,tail=1;queue[0]=start;seen[start]=1;let x0=w,x1=0,y0=h,y1=0,count=0,sx=0,sy=0;
      while(head<tail){const n=queue[head++],x=n%w,y=(n/w)|0;count++;sx+=x;sy+=y;x0=Math.min(x0,x);x1=Math.max(x1,x);y0=Math.min(y0,y);y1=Math.max(y1,y);
        if(x>0&&!seen[n-1]&&g[n-1]<93){seen[n-1]=1;queue[tail++]=n-1;}
        if(x<w-1&&!seen[n+1]&&g[n+1]<93){seen[n+1]=1;queue[tail++]=n+1;}
        if(y>0&&!seen[n-w]&&g[n-w]<93){seen[n-w]=1;queue[tail++]=n-w;}
        if(y<h-1&&!seen[n+w]&&g[n+w]<93){seen[n+w]=1;queue[tail++]=n+w;}
      }
      const bw=x1-x0+1,bh=y1-y0+1,fill=count/(bw*bh);
      if(bw>=min&&bh>=min&&bw<=max&&bh<=max&&bw/bh>.65&&bw/bh<1.55&&fill>.84)candidates.push({x:sx/count,y:sy/count,size:Math.sqrt(count),fill});
    }
    if(candidates.length<4)return null;
    // Choose a geometrically consistent set, rather than trusting a text glyph.
    const quadrants=[candidates.filter(p=>p.x<w*.5&&p.y<h*.55),candidates.filter(p=>p.x>w*.5&&p.y<h*.55),candidates.filter(p=>p.x>w*.5&&p.y>h*.55),candidates.filter(p=>p.x<w*.5&&p.y>h*.55)];
    if(quadrants.some(a=>!a.length))return null;
    quadrants.forEach(a=>a.sort((a,b)=>b.size-a.size));let best=null,bestScore=-1;
    for(const a of quadrants[0].slice(0,8))for(const b of quadrants[1].slice(0,8))for(const c of quadrants[2].slice(0,8))for(const d of quadrants[3].slice(0,8)){
      const p=[a,b,c,d];if(!validCorners(p,w,h))continue;const sizes=p.map(p=>p.size);if(Math.min(...sizes)/Math.max(...sizes)<.55)continue;
      const width=(Math.hypot(a.x-b.x,a.y-b.y)+Math.hypot(c.x-d.x,c.y-d.y))/2,height=(Math.hypot(a.x-d.x,a.y-d.y)+Math.hypot(c.x-b.x,c.y-b.y))/2;
      if(width/height<.48||width/height>1.3||Math.abs(a.y-b.y)>height*.24||Math.abs(c.y-d.y)>height*.24)continue;
      const score=width*height*(Math.min(...sizes)/Math.max(...sizes));if(score>bestScore){bestScore=score;best=p.map(({x,y})=>({x,y}));}
    }return best;
  }
  function classify(scores){
    const ordered=scores.map((v,i)=>({v,i})).sort((a,b)=>b.v-a.v),first=ordered[0],second=ordered[1];
    if(first.v<.18)return {answer:null,status:'blank',confidence:1-first.v};
    if(second.v>=.4)return {answer:null,status:'multiple',confidence:0};
    if(first.v<.48||first.v-second.v<.25)return {answer:null,status:'uncertain',confidence:first.v-second.v};
    return {answer:first.i,status:'clear',confidence:Math.max(0,Math.min(1,first.v-second.v))};
  }
  function analyze(image,count,options,corners){
    if(!validCorners(corners,image.width,image.height))throw Error('المربعات غير مكتملة أو ترتيبها غير صحيح. أعيدي تحديدها.');
    const h=homography(MARKERS,corners),g=gray(image),w=image.width,ht=image.height;
    const read=(x,y)=>{const p=project(h,x,y),xx=Math.round(p.x),yy=Math.round(p.y);if(xx<0||yy<0||xx>=w||yy>=ht)throw Error('بعض الإجابات خارج الصورة. أظهري الورقة كاملة.');return g[yy*w+xx];};
    // An asymmetric short bar beside the top-left square prevents an upside-down
    // sheet from being accepted as a valid but incorrectly mapped answer sheet.
    const orientationBg=read(93,190);let orientationDark=0;
    for(let yy=-1;yy<=1;yy++)for(let xx=-3;xx<=3;xx++)orientationDark+=read(71+xx,190+yy)<Math.max(50,orientationBg*.65)?1:0;
    if(orientationDark<15)throw Error('علامة اتجاه النموذج غير واضحة. تأكد من استخدام نموذج بصيرة، وأن عنوان الورقة في الأعلى؛ استخدم زر التدوير عند الحاجة.');
    return layout(count,options).map(row=>{
      const scores=row.bubbles.map(({x,y})=>{
        const bg=[];for(let i=0;i<24;i++){const t=i*Math.PI/12;bg.push(read(x+Math.cos(t)*row.r*1.45,y+Math.sin(t)*row.r*1.45));}bg.sort((a,b)=>a-b);const white=bg[18];
        let dark=0,total=0;for(let yy=-5;yy<=5;yy++)for(let xx=-5;xx<=5;xx++){if(xx*xx+yy*yy>25)continue;const v=read(x+xx*row.r*.12,y+yy*row.r*.12);dark+=v<Math.max(50,white*.65)?1:0;total++;}return dark/total;
      });return {...classify(scores),scores};
    });
  }
  function grade(quiz,answers){let score=0,total=0,correct=0;quiz.questions.forEach((q,i)=>{total+=q.weight;if(answers[i]===q.answer){score+=q.weight;correct++;}});return {score,total,correct,percent:total?score/total*100:0};}
  function aggregate(quiz,results){
    const rows=results.map(r=>({...r,...grade(quiz,r.answers)}));
    const items=quiz.questions.map((q,i)=>{const counts=Array(quiz.options).fill(0);let blank=0;rows.forEach(r=>r.answers[i]===null?blank++:counts[r.answers[i]]++);return {number:i+1,skill:q.skill||'غير مصنّف',correct:counts[q.answer],blank,counts,percent:rows.length?counts[q.answer]/rows.length*100:0};});
    const skills=Object.create(null);quiz.questions.forEach((q,i)=>{const name=q.skill||'غير مصنّف';skills[name]??={name,earned:0,total:0,questions:0};skills[name].earned+=items[i].correct*q.weight;skills[name].total+=rows.length*q.weight;skills[name].questions++;});
    return {rows,items,skills:Object.values(skills).map(s=>({...s,percent:s.total?s.earned/s.total*100:0})),mean:rows.length?rows.reduce((sum,r)=>sum+r.percent,0)/rows.length:0};
  }
  function csvCell(value){let s=String(value??'');if(/^[\s]*[=+@\-\t\r]/.test(s))s="'"+s;return '"'+s.replace(/"/g,'""')+'"';}
  return {WIDTH,HEIGHT,MARKERS,layout,homography,project,validCorners,detectMarkers,classify,analyze,grade,aggregate,csvCell};
});
