/* ATME creative layer: constellation canvas, spotlight, 3D tilt, typed text, theme switch, mapping graph */
(function(){
  var reduce=window.matchMedia&&window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  var touch=window.matchMedia&&window.matchMedia('(hover: none)').matches;
  var root=document.documentElement;

  /* ---------- theme ---------- */
  root.setAttribute('data-theme','dark');

  /* ---------- cursor spotlight ---------- */
  if(!touch&&!reduce){
    var sp=document.createElement('div');sp.className='spot';document.body.appendChild(sp);
    window.addEventListener('mousemove',function(e){sp.style.transform='translate('+(e.clientX-260)+'px,'+(e.clientY-260)+'px)';sp.style.opacity=1;},{passive:true});
  }

  /* ---------- constellation canvas ---------- */
  function constellation(host){
    var c=document.createElement('canvas');c.className='const-canvas';c.setAttribute('aria-hidden','true');host.insertBefore(c,host.firstChild);
    var ctx=c.getContext('2d'),W=0,H=0,dpr=Math.min(window.devicePixelRatio||1,2),nodes=[],mouse={x:-999,y:-999};
    var colsD=['#6BDCD6','#F2A93B','#6EE7A8','#FFFFFF'],colsL=['#3BB8B0','#F2A93B','#5BBDB8','#8CD2CE'];
    function size(){W=host.clientWidth;H=host.clientHeight;c.width=W*dpr;c.height=H*dpr;ctx.setTransform(dpr,0,0,dpr,0,0);
      var n=Math.max(18,Math.min(70,Math.round(W*H/11000)));nodes=[];
      for(var i=0;i<n;i++)nodes.push({x:Math.random()*W,y:Math.random()*H,vx:(Math.random()-.5)*.35,vy:(Math.random()-.5)*.35,r:1.6+Math.random()*2.4,i:i});}
    function draw(){
      ctx.clearRect(0,0,W,H);var light=root.getAttribute('data-theme')!=='dark',pal=light?colsL:colsD,lc=light?'15,124,123':'255,255,255',k0=light?.34:.22;
      for(var i=0;i<nodes.length;i++){var a=nodes[i];
        if(!reduce){a.x+=a.vx;a.y+=a.vy;if(a.x<0||a.x>W)a.vx*=-1;if(a.y<0||a.y>H)a.vy*=-1;
          var dx=mouse.x-a.x,dy=mouse.y-a.y,d=Math.sqrt(dx*dx+dy*dy);if(d<140&&d>1){a.x+=dx/d*.6;a.y+=dy/d*.6;}}
        for(var j=i+1;j<nodes.length;j++){var b=nodes[j],ex=a.x-b.x,ey=a.y-b.y,dd=ex*ex+ey*ey;
          if(dd<15000){ctx.strokeStyle='rgba('+lc+','+(k0*(1-dd/15000))+')';ctx.lineWidth=1;ctx.beginPath();ctx.moveTo(a.x,a.y);ctx.lineTo(b.x,b.y);ctx.stroke();}}
        ctx.fillStyle=pal[a.i%pal.length];ctx.globalAlpha=.9;ctx.beginPath();ctx.arc(a.x,a.y,a.r,0,6.283);ctx.fill();ctx.globalAlpha=1;}
      // mouse links
      for(var k=0;k<nodes.length;k++){var m=nodes[k],mx=m.x-mouse.x,my=m.y-mouse.y,md=mx*mx+my*my;
        if(md<22000){ctx.strokeStyle='rgba(242,169,59,'+(0.55*(1-md/22000))+')';ctx.beginPath();ctx.moveTo(m.x,m.y);ctx.lineTo(mouse.x,mouse.y);ctx.stroke();}}
      if(!reduce&&!document.hidden)requestAnimationFrame(draw);else if(!reduce)setTimeout(function(){requestAnimationFrame(draw)},400);
    }
    host.addEventListener('mousemove',function(e){var r=host.getBoundingClientRect();mouse.x=e.clientX-r.left;mouse.y=e.clientY-r.top;});
    host.addEventListener('mouseleave',function(){mouse.x=mouse.y=-999;});
    window.addEventListener('resize',size);window.addEventListener('load',size);
    if(window.ResizeObserver){new ResizeObserver(function(){if(Math.abs(host.clientHeight-H)>4||Math.abs(host.clientWidth-W)>4)size();}).observe(host);}
    size();draw();
  }
  document.querySelectorAll('[data-constellation]').forEach(constellation);

  /* ---------- 3D tilt ---------- */
  if(!touch&&!reduce){
    document.querySelectorAll('.cards .card,.kpi-card,.co-card,.drop,.ai-card,.export-btn,.analytics-card,.step').forEach(function(el){
      el.addEventListener('mousemove',function(e){var r=el.getBoundingClientRect(),x=(e.clientX-r.left)/r.width-.5,y=(e.clientY-r.top)/r.height-.5;
        el.style.transition='transform .08s';el.style.transform='perspective(700px) rotateX('+(-y*7)+'deg) rotateY('+(x*9)+'deg) translateY(-5px)';});
      el.addEventListener('mouseleave',function(){el.style.transition='transform .4s';el.style.transform='';});
    });
  }

  /* ---------- typed text ---------- */
  document.querySelectorAll('[data-typed]').forEach(function(el){
    var words=el.getAttribute('data-typed').split('|'),wi=0,ci=0,del=false;
    if(reduce){el.textContent=words[0];return;}
    (function tick(){var w=words[wi];el.textContent=w.slice(0,ci);
      if(!del&&ci<w.length){ci++;setTimeout(tick,70);}
      else if(!del){del=true;setTimeout(tick,1400);}
      else if(ci>0){ci--;setTimeout(tick,35);}
      else{del=false;wi=(wi+1)%words.length;setTimeout(tick,250);}})();
  });


  /* ---------- charts: theme + bar/pie switch ---------- */
  var PAL=['#238987','#F2A93B','#1A6B69','#5BBDB8','#C98412','#8CD2CE','#2EA3A0','#F9CDA8','#0F4F4D','#A9D9D5','#14706E','#E08A3C'];
  function chartTheme(){
    if(!window.Chart)return;var dark=root.getAttribute('data-theme')==='dark';
    var tick=dark?'#D3E8E5':'#587A78',grid=dark?'rgba(255,255,255,.16)':'rgba(31,68,66,.09)',leg=dark?'#D3E8E5':'#1F3F3D';
    Chart.defaults.color=tick;Chart.defaults.borderColor=grid;
    Object.keys(Chart.instances||{}).forEach(function(k){var c=Chart.instances[k],o=c.options;
      if(o.scales)Object.keys(o.scales).forEach(function(n){var sc=o.scales[n];sc.grid=sc.grid||{};if(sc.grid.display!==false)sc.grid.color=grid;sc.ticks=sc.ticks||{};sc.ticks.color=tick;});
      if(o.plugins){o.plugins.legend=o.plugins.legend||{};o.plugins.legend.labels=o.plugins.legend.labels||{};o.plugins.legend.labels.color=leg;}
      c.update('none');});
  }
  function addSwitches(){
    if(!window.Chart||!Chart.getChart)return;
    document.querySelectorAll('canvas').forEach(function(cv){
      var ch=Chart.getChart(cv);if(!ch||ch.config.type!=='bar')return;
      var host=cv.closest('.chart-card')||cv.closest('.box');if(!host||host.querySelector('.ct-switch'))return;
      host.style.position='relative';
      var sw=document.createElement('div');sw.className='ct-switch';sw.innerHTML='<button type="button" class="on" data-t="bar">Bar</button><button type="button" data-t="pie">Pie</button>';host.appendChild(sw);
      var orig={data:ch.config.data,options:ch.config.options};
      sw.addEventListener('click',function(e){var b=e.target.closest('button');if(!b)return;
        sw.querySelectorAll('button').forEach(function(x){x.classList.toggle('on',x===b);});
        var cur=Chart.getChart(cv);if(cur)cur.destroy();
        if(b.getAttribute('data-t')==='bar'){new Chart(cv,{type:'bar',data:orig.data,options:orig.options});}
        else{var ds=orig.data.datasets[0],n=orig.data.labels.length,cols=[];for(var i=0;i<n;i++)cols.push(PAL[i%PAL.length]);
          new Chart(cv,{type:'pie',data:{labels:orig.data.labels,datasets:[{label:ds.label,data:ds.data,backgroundColor:cols,borderWidth:0,spacing:3,hoverOffset:14}]},
            options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:n>6?'right':'bottom',labels:{boxWidth:12,padding:10}}}}});}
        chartTheme();});
    });
  }
  window.addEventListener('load',function(){addSwitches();chartTheme();});

  /* ---------- report: interactive mapping graph ---------- */
  var host=document.getElementById('constellation');
  if(host){
    function readTable(id){var sec=document.getElementById(id);if(!sec)return null;var t=sec.querySelector('table.matrix');if(!t)return null;
      var labels=[].slice.call(t.querySelectorAll('thead th')).slice(1).map(function(x){return x.textContent.trim();});
      var rows=[].slice.call(t.querySelectorAll('tbody tr')).map(function(tr){var tds=tr.querySelectorAll('td');
        return {co:tds[0].textContent.trim(),v:[].slice.call(tds).slice(1).map(function(td){var n=parseInt(td.textContent,10);return isNaN(n)?0:n;})};});
      return {labels:labels,rows:rows};}
    var sets={PO:readTable('co-po-matrix'),PEO:readTable('co-peo-matrix'),PSO:readTable('co-pso-matrix')};
    var tabs=host.querySelector('.cg-tabs'),stage=host.querySelector('.cg-stage'),cur=null;
    var col={3:'#1E8E5A',2:'#238987',1:'#F2A93B'};
    function render(key){
      cur=key;var d=sets[key];if(!d){stage.innerHTML='<p class="cg-empty">No data</p>';return;}
      var W=Math.max(stage.clientWidth,520),nR=d.rows.length,nC=d.labels.length,H=Math.max(nR,nC)*44+40,xl=90,xr=W-90;
      var yl=function(i){return 28+(H-56)*(nR===1?.5:i/(nR-1));},yr=function(j){return 28+(H-56)*(nC===1?.5:j/(nC-1));};
      var s='<svg viewBox="0 0 '+W+' '+H+'" width="100%" height="'+H+'" role="img" aria-label="CO to '+key+' mapping graph">';
      d.rows.forEach(function(r,i){r.v.forEach(function(v,j){if(!v)return;var x1=xl,y1=yl(i),x2=xr,y2=yr(j),mx=(x1+x2)/2;
        s+='<path class="cg-link" data-co="'+i+'" data-t="'+j+'" d="M'+x1+' '+y1+' C'+mx+' '+y1+' '+mx+' '+y2+' '+x2+' '+y2+'" stroke="'+col[v]+'" stroke-width="'+(v*1.3+.6)+'" fill="none" opacity=".55"/>';});});
      d.rows.forEach(function(r,i){s+='<g class="cg-node co" data-co="'+i+'"><circle cx="'+xl+'" cy="'+yl(i)+'" r="15"/><text x="'+(xl-24)+'" y="'+(yl(i)+4)+'" text-anchor="end">'+r.co+'</text></g>';});
      d.labels.forEach(function(l,j){s+='<g class="cg-node tg" data-t="'+j+'"><circle cx="'+xr+'" cy="'+yr(j)+'" r="13"/><text x="'+(xr+24)+'" y="'+(yr(j)+4)+'">'+l+'</text></g>';});
      stage.innerHTML=s+'</svg>';
      var links=stage.querySelectorAll('.cg-link');
      function focus(attr,val){links.forEach(function(p){var on=p.getAttribute(attr)===val;p.style.opacity=on?1:.06;p.style.strokeWidth=on?(parseFloat(p.getAttribute('stroke-width'))+1.5):'';});}
      function reset(){links.forEach(function(p){p.style.opacity='';p.style.strokeWidth='';});}
      stage.querySelectorAll('.cg-node.co').forEach(function(n){n.addEventListener('mouseenter',function(){focus('data-co',n.getAttribute('data-co'));});n.addEventListener('mouseleave',reset);});
      stage.querySelectorAll('.cg-node.tg').forEach(function(n){n.addEventListener('mouseenter',function(){focus('data-t',n.getAttribute('data-t'));});n.addEventListener('mouseleave',reset);});
    }
    Object.keys(sets).forEach(function(k,idx){if(!sets[k])return;var b=document.createElement('button');b.type='button';b.textContent='CO → '+k;b.setAttribute('data-k',k);
      b.onclick=function(){tabs.querySelectorAll('button').forEach(function(x){x.classList.remove('on');});b.classList.add('on');render(k);};tabs.appendChild(b);});
    var first=tabs.querySelector('button');if(first){first.classList.add('on');render(first.getAttribute('data-k'));}else{host.style.display='none';}
    window.addEventListener('resize',function(){if(cur)render(cur);});
  }
})();
