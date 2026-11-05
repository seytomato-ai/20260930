(function(){
  'use strict';
  const teachers = v => Clubs.teachers(v);
  // Group before merging periods. Unknown teachers remain separate to avoid losing unresolved requests.
  function unique(rows) {
    const groups=[];
    for(const original of rows) {
      const r={...original}, names=teachers(r.teacher);
      const matches=names.length?groups.filter(g=>g.date===r.date && g.periodText===r.periodText && teachers(g.teacher).some(t=>names.includes(t))):[];
      if(!matches.length) {r.teacher=names.join(', ');groups.push(r);continue;}
      const all=[...matches,r], target=matches[0];
      const distinct=key=>[...new Set(all.flatMap(x=>String(x[key]||'').split(', ').filter(Boolean)))].join(', ');
      target.name=distinct('name'); target.teacher=teachers(all.map(x=>x.teacher).join(',')).join(', ');
      target.subject=distinct('subject'); target.classCode=distinct('classCode');
      target.groupTotal=target.name.split(', ').length;
      target.missing=all.some(x=>x.missing);
      for(const g of matches.slice(1)) groups.splice(groups.indexOf(g),1);
    }
    return groups;
  }
  window.RequestRows={unique};
})();
