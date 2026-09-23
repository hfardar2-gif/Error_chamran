/* Office charts stay editable. Photographs come from the user's reference deck. */
(function(){
const W=13.333,H=7.5;
const P={navy:'103A4C',ink:'173543',teal:'0C918C',cyan:'D8F2F0',blue:'3C78BF',gold:'E4A44E',gray:'627C8A',rule:'D6E4E8',white:'FFFFFF',pale:'F5F9FA'};
const font='Vazirmatn';
const iconPaths={
 activity:'<path d="M2 12h4l3-7 5 14 3-7h5"/>',
 shield:'<path d="M12 2 20 6v6c0 5-3 8-8 10C7 20 4 17 4 12V6z"/><path d="m9 12 2 2 4-4"/>',
 building:'<rect x="4" y="3" width="16" height="18" rx="2"/><path d="M9 8h6M12 5v6M10 21v-5h4v5"/>',
 clipboard:'<rect x="5" y="4" width="14" height="18" rx="2"/><rect x="9" y="2" width="6" height="4" rx="1"/><path d="M8 11h8M8 15h8"/>',
 calendar:'<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M7 2v6M17 2v6M3 10h18"/>',
 bars:'<path d="M4 20V11h4v9M10 20V5h4v15M16 20v-7h4v7M2 20h20"/>',
 clock:'<circle cx="12" cy="12" r="9"/><path d="M12 7v5l4 2"/>',
 people:'<circle cx="8" cy="8" r="3"/><circle cx="17" cy="9" r="2.5"/><path d="M2 20v-2a6 6 0 0 1 12 0v2M15 15a5 5 0 0 1 7 5"/>',
 pill:'<path d="M5 20a5 5 0 0 1-1-7L13 4a5 5 0 0 1 7 7l-9 9a5 5 0 0 1-6 0Z"/><path d="m8 9 7 7"/>',
 alert:'<path d="m12 3 10 18H2L12 3Z"/><path d="M12 9v5M12 18h.01"/>',
 search:'<circle cx="10" cy="10" r="7"/><path d="m15 15 6 6"/>',
 layers:'<path d="m12 3 10 5-10 5L2 8l10-5ZM2 13l10 5 10-5M2 18l10 5 10-5"/>',
 grid:'<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18M15 3v18"/>'
};
const map={shift:'clock',trend:'activity',occurrence:'building',reporting:'clipboard',cross:'grid',reportType:'clipboard',severity:'shield',errorKind:'alert',causes:'search',roles:'people',age:'people',mainType:'layers',medicineShift:'pill',missing:'grid'};
function iconSvg(name,color='#0C918C'){
 const paths=iconPaths[map[name]||name]||iconPaths.activity;
 return `<svg xmlns="http://www.w3.org/2000/svg" width="80" height="80" viewBox="0 0 24 24" fill="none" stroke="${color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">${paths}</svg>`;
}
function iconData(name,color){return 'data:image/svg+xml;base64,'+btoa(iconSvg(name,color))}
window.chartIcon=(name)=>iconSvg(name);
function shortLabel(label){let s=String(label).replace(/\([^)]*\)/g,'').replace(/\s+/g,' ').trim();return s.length>25?s.slice(0,24)+'…':s}
function addText(s,text,box,extra={}){s.addText(String(text),{...box,fontFace:font,margin:0,color:P.ink,rtlMode:true,align:'right',valign:'mid',...extra})}
function addIcon(s,pptx,id,x,y,size=.38,color='#0C918C'){
 try{s.addImage({data:iconData(id,color),x,y,w:size,h:size})}
 catch(e){s.addShape(pptx.ShapeType.ellipse,{x,y,w:size,h:size,line:{color:P.teal},fill:{color:P.white}})}
}
function base(pptx,section){const s=pptx.addSlide();s.background={color:P.white};s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:.12,h:H,line:{color:P.teal},fill:{color:P.teal}});s.addShape(pptx.ShapeType.line,{x:.68,y:7.07,w:12.05,h:0,line:{color:P.rule,width:.7}});addText(s,'واحد ایمنی بیمار  |  '+section,{x:7.2,y:7.13,w:5.35,h:.19},{fontSize:8,color:P.gray});return s}
function title(s,pptx,label,id,eyebrow){addIcon(s,pptx,id,11.9,.49,.48);addText(s,label,{x:.7,y:.56,w:10.95,h:.58},{fontSize:24,bold:true,color:P.navy});if(eyebrow)addText(s,eyebrow,{x:.7,y:1.22,w:11.67,h:.28},{fontSize:10,color:P.gray});s.addShape(pptx.ShapeType.line,{x:.7,y:1.6,w:11.95,h:0,line:{color:P.rule,width:1}})}
function photo(s,data,x,y,w,h){if(data)s.addImage({data,x,y,w,h})}
function cover(pptx,state,assets){const s=pptx.addSlide();s.background={color:P.white};photo(s,assets['hospital-corridor'],0,0,4.5,H);s.addShape(pptx.ShapeType.rect,{x:4.5,y:0,w:.14,h:H,line:{color:P.teal},fill:{color:P.teal}});addIcon(s,pptx,'building',11.76,.8,.72);addText(s,'گزارش پایش خطاهای ثبت شده',{x:5.2,y:1.76,w:7.15,h:1.15},{fontSize:31,bold:true,color:P.navy});addText(s,'بیمارستان قلب شهید چمران اصفهان',{x:5.2,y:3.02,w:7.15,h:.55},{fontSize:17,color:P.teal});s.addShape(pptx.ShapeType.line,{x:6.7,y:4.12,w:5.6,h:0,line:{color:P.rule,width:1.2}});addText(s,`سال غالبِ تاریخ وقوع: ${state.period}`,{x:5.2,y:4.4,w:7.15,h:.44},{fontSize:15,bold:true});addText(s,`${state.rows.length} گزارش ثبت شده  |  ${state.trendRows.length} تاریخ در سال غالب`,{x:5.2,y:5.15,w:7.15,h:.42},{fontSize:12,color:P.gray});addText(s,'واحد ایمنی بیمار',{x:5.2,y:6.82,w:7.15,h:.25},{fontSize:10,color:P.gray});return s}
function overview(pptx,state,assets){const s=base(pptx,'خلاصه داده');photo(s,assets['patient-safety-dashboard'],.12,.12,3.75,6.93);s.addShape(pptx.ShapeType.rect,{x:3.88,y:.12,w:.11,h:6.93,line:{color:P.teal},fill:{color:P.teal}});addText(s,'دادهٔ ورودی، پیش از تفسیر',{x:4.3,y:.67,w:8.3,h:.65},{fontSize:25,bold:true,color:P.navy});const stats=[[state.rows.length,'همه گزارش‌ها'],[state.trendRows.length,'دارای تاریخ در سال غالب'],[state.errors,'بدون تاریخ معتبر'],[state.outside,'تاریخ مربوط به سال دیگر']];stats.forEach(([v,l],i)=>{let x=4.32+i%2*4.12,y=1.75+Math.floor(i/2)*1.7;s.addShape(pptx.ShapeType.line,{x,y:y+1.18,w:3.7,h:0,line:{color:P.rule,width:1}});addText(s,v,{x,y,w:3.7,h:.72},{fontSize:34,bold:true,color:i>1?P.gold:P.teal});addText(s,l,{x,y:y+.75,w:3.7,h:.35},{fontSize:12,bold:true,color:P.gray})});addText(s,'همه رکوردها در نمودارهای دسته‌ای شمرده می‌شوند. نمودار روند فقط تاریخ‌های معتبر سال غالب را نشان می‌دهد.',{x:4.3,y:5.5,w:8.15,h:.78},{fontSize:12,color:P.ink});return s}
function chartSeries(c){if(c.type==='cross'){const labels=c.data.map(d=>shortLabel(d[0]));return (c.series||[]).map((shift,i)=>({name:shift,labels,values:c.data.map(d=>Number(d[i+1])||0)}))}return [{name:'تعداد گزارش',labels:c.data.map(d=>shortLabel(d[0])),values:c.data.map(d=>Number(d[1])||0)}]}
function chartSlide(pptx,c,state){const s=base(pptx,'نمودار');title(s,pptx,c.title,c.id,c.subtitle);const series=chartSeries(c);const kind=c.type==='line'?pptx.ChartType.line:c.type==='donut'?pptx.ChartType.pie:pptx.ChartType.bar;
 const bars=c.type==='bar'||c.type==='cross';const options={x:1.03,y:1.85,w:11.15,h:4.62,chartColors:[P.teal,P.blue,P.gold,'9B79B8','D8756F'],showTitle:false,showLegend:c.type==='cross'||c.type==='donut',legendPos:'b',legendFontFace:font,legendFontSize:10,showValue:c.type!=='cross',dataLabelFontFace:font,dataLabelFontBold:true,dataLabelFontSize:11,dataLabelColor:P.ink,dataLabelPosition:bars?'outEnd':'bestFit',showMarker:c.type==='line',showLine:true,showBorder:false,catAxisLabelFontFace:font,catAxisLabelFontBold:true,catAxisLabelFontSize:bars?10:11,catAxisLabelColor:P.ink,valAxisLabelFontFace:font,valAxisLabelFontBold:true,valAxisLabelFontSize:10,valAxisMinVal:0,valGridLine:{color:P.rule,width:.6},showShadow:false,barDir:'bar',barGrouping:c.type==='cross'?'stacked':'clustered',showHoleSize:55,layout:{x:.10,y:.08,w:.82,h:.8}};
 s.addChart(kind,series,options);
 const coverage=c.id==='trend'?`${state.trendRows.length} رکورد با تاریخ معتبر در سال غالب؛ ${state.rows.length-state.trendRows.length} رکورد خارج از این نمودار زمانی.`:c.id==='cross'?'این نمودار پنج واحد پرتکرار گزارش‌دهنده را در همه شیفت‌ها نشان می‌دهد.':`مبنای محاسبه: ${state.rows.length} گزارش ثبت شده؛ «ثبت نشده» در گروه‌بندی حفظ شده است.`;
 addText(s,coverage,{x:.9,y:6.62,w:11.5,h:.27},{fontSize:9,color:P.gray});return s}
function analysisSlide(pptx,c,state,analysisFor){const s=base(pptx,'تحلیل و اقدام');title(s,pptx,'تحلیل | '+c.title,c.id,'مبنای تحلیل: گزارش‌های ثبت شده، بدون ادعای علت یا نرخ خطر');const a=analysisFor(c),sections=[['۱','یافتهٔ قابل اتکا',a.observation],['۲','پرسش پیش از نتیجه‌گیری',a.question],['۳','اقدام برای بررسی',a.action]];
 sections.forEach(([num,label,body],i)=>{let y=1.93+i*1.32;addText(s,num,{x:.8,y,w:.55,h:.65},{fontSize:30,bold:true,color:P.teal,align:'center'});addText(s,label,{x:1.62,y:y+.02,w:10.75,h:.33},{fontSize:14,bold:true,color:P.navy});addText(s,body,{x:1.62,y:y+.44,w:10.75,h:.58},{fontSize:14,color:P.ink});s.addShape(pptx.ShapeType.line,{x:1.62,y:y+1.13,w:10.75,h:0,line:{color:P.rule,width:.7}})});
 if(c.note)addText(s,'تصمیم ثبت شده: '+c.note.slice(0,220),{x:1.62,y:6.02,w:10.75,h:.38},{fontSize:11,bold:true,color:P.navy});addText(s,a.caveat,{x:1.62,y:6.57,w:10.75,h:.3},{fontSize:9,color:P.gray});return s}
function closing(pptx,state,assets){const s=pptx.addSlide();s.background={color:P.navy};photo(s,assets['clinical-review'],0,0,4.1,H);s.addShape(pptx.ShapeType.rect,{x:4.1,y:0,w:.13,h:H,line:{color:P.teal},fill:{color:P.teal}});addIcon(s,pptx,'shield',11.85,.85,.69,'#FFFFFF');addText(s,'از ثبت خطا تا تصمیم قابل پیگیری',{x:4.7,y:1.77,w:7.85,h:1.15},{fontSize:28,bold:true,color:P.white});addText(s,'۱. داده و پوشش هر نمودار را تأیید کنید.\n۲. موارد پرتکرار را با پرونده و حجم فعالیت بسنجید.\n۳. اقدام، مسئول و موعد بازبینی را در کمیته ثبت کنید.',{x:4.75,y:3.16,w:7.5,h:2.12},{fontSize:16,color:'D9EDEB',breakLine:false});addText(s,'واحد ایمنی بیمار | '+state.period,{x:4.7,y:6.83,w:7.8,h:.27},{fontSize:10,color:'AFD6D3'});return s}
window.makeSafetyPresentation=async function(state,analysisFor){const pptx=new PptxGenJS();pptx.layout='LAYOUT_WIDE';pptx.author='واحد ایمنی بیمار';pptx.subject='گزارش آماری خطاهای ثبت شده';pptx.title='گزارش ایمنی بیمار · '+state.period;pptx.lang='fa-IR';pptx.rtlMode=true;const assets=window.REPORT_ASSETS||{};cover(pptx,state,assets);overview(pptx,state,assets);for(const c of state.charts.filter(c=>c.selected)){chartSlide(pptx,c,state);analysisSlide(pptx,c,state,analysisFor)}closing(pptx,state,assets);await pptx.writeFile({fileName:`گزارش-ایمنی-بیمار-${state.period.replace(/[^\d\p{L}]+/gu,'-')}.pptx`})};
})();
