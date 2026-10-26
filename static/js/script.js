/* ATME UI helpers: count-up numbers, scroll reveal, scroll-spy, back-to-top, button ripple. */
(function(){
  var reduce=window.matchMedia&&window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  // count-up for elements with data-count (or .card h2 / .kpi-value that are pure integers)
  function countUp(el){
    var raw=(el.textContent||'').trim();
    if(!/^\d+$/.test(raw)||reduce)return;
    var end=parseInt(raw,10),t0=null,dur=900;
    function step(t){if(!t0)t0=t;var p=Math.min((t-t0)/dur,1),e=1-Math.pow(1-p,3);
      el.textContent=Math.round(end*e);if(p<1)requestAnimationFrame(step);else el.textContent=end;}
    el.textContent='0';requestAnimationFrame(step);
  }
  document.querySelectorAll('.card h2,.kpi-value,.analytics-card h1').forEach(countUp);

  // scroll reveal
  var rev=document.querySelectorAll('.box,.card2,.co-card,.ai-grid>*,.export-btn');
  if('IntersectionObserver' in window&&!reduce){
    rev.forEach(function(e){e.classList.add('reveal');});
    var io=new IntersectionObserver(function(es){es.forEach(function(x){if(x.isIntersecting){x.target.classList.add('in');io.unobserve(x.target);}});},{threshold:.08});
    rev.forEach(function(e){io.observe(e);});
  }

  // scroll-spy for sidebar links and report pill-nav
  var links=[].slice.call(document.querySelectorAll('.sidebar a[href^="#"],.pill-nav a[href^="#"]'));
  if(links.length&&'IntersectionObserver' in window){
    var map={};links.forEach(function(a){var id=a.getAttribute('href').slice(1);var s=document.getElementById(id);if(s)map[id]=a;});
    var spy=new IntersectionObserver(function(es){es.forEach(function(x){
      if(x.isIntersecting){links.forEach(function(a){a.classList.remove('spy');});
        var a=map[x.target.id];if(a){a.classList.add('spy');if(a.closest('.pill-nav'))a.scrollIntoView({block:'nearest',inline:'center'});}}});},{rootMargin:'-35% 0px -60% 0px'});
    Object.keys(map).forEach(function(id){spy.observe(document.getElementById(id));});
  }

  // back-to-top
  var b=document.createElement('button');b.className='to-top';b.setAttribute('aria-label','Back to top');b.innerHTML='&uarr;';
  b.onclick=function(){window.scrollTo({top:0,behavior:'smooth'});};document.body.appendChild(b);
  window.addEventListener('scroll',function(){b.classList.toggle('show',window.scrollY>500);},{passive:true});

  // ripple on buttons
  document.addEventListener('click',function(e){
    var t=e.target.closest('.primary-btn,.go,.box button,.nav-btn,.export-btn');if(!t||reduce)return;
    var r=document.createElement('span');r.className='ripple';var rc=t.getBoundingClientRect(),d=Math.max(rc.width,rc.height);
    r.style.cssText='width:'+d+'px;height:'+d+'px;left:'+(e.clientX-rc.left-d/2)+'px;top:'+(e.clientY-rc.top-d/2)+'px';
    if(getComputedStyle(t).position==='static')t.style.position='relative';t.style.overflow='hidden';t.appendChild(r);setTimeout(function(){r.remove();},600);
  });
})();
