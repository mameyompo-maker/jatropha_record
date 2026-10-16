const fs=require('fs');
const path=require('path');
const HERE=__dirname;
const OUT=path.join(HERE,'..','Seed_Storage_Sowing_Plan_2026_EN.docx');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,ShadingType,AlignmentType,HeadingLevel,ImageRun,LevelFormat,BorderStyle,PageBreak}=require('docx');
const P=JSON.parse(fs.readFileSync(path.join(HERE,'plan.json')));
const WD=['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];
const MON=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const mdw=(s,w)=>WD[w]+' '+md(s);
const md=s=>{if(!s||s==='None')return '—';const [y,m,d]=s.split('-');return `${+d} ${MON[+m-1]}`};
const n=v=>v.toLocaleString('en-US');
const FONT={ascii:'Arial',hAnsi:'Arial',eastAsia:'Arial'};
const t=(s,o={})=>new TextRun({text:s,font:FONT,...o});
const p=(s,o={})=>new Paragraph({children:Array.isArray(s)?s:[t(s)],spacing:{after:100},...o});
const h=(s,l)=>new Paragraph({heading:l,keepNext:true,children:[t(s)],spacing:{before:240,after:120}});
const b=(s)=>new Paragraph({numbering:{reference:'bul',level:0},children:Array.isArray(s)?s:[t(s)],spacing:{after:60}});
const border={style:BorderStyle.SINGLE,size:4,color:'BBBBBB'};
const borders={top:border,bottom:border,left:border,right:border};
function table(head,rows,widths,opts={}){
  const total=widths.reduce((a,c)=>a+c,0);
  const cell=(s,w,hd,fill)=>new TableCell({width:{size:w,type:WidthType.DXA},borders,
    shading:hd?{type:ShadingType.CLEAR,fill:'1F4E79',color:'auto'}:(fill?{type:ShadingType.CLEAR,fill,color:'auto'}:undefined),
    margins:{top:40,bottom:40,left:80,right:80},
    children:[new Paragraph({alignment:hd?AlignmentType.CENTER:(/^[\d,.\-—+ kg%/→]+$/.test(s)?AlignmentType.RIGHT:AlignmentType.LEFT),
      children:[t(s,{bold:!!hd,color:hd?'FFFFFF':undefined,size:opts.size||18})]})]});
  return new Table({width:{size:total,type:WidthType.DXA},columnWidths:widths,
    rows:[new TableRow({tableHeader:true,children:head.map((s,i)=>cell(s,widths[i],true))}),
      ...rows.map((r,ri)=>new TableRow({children:r.map((s,i)=>cell(String(s),widths[i],false,opts.fill?opts.fill(r,ri):undefined))}))]});
}
const L=P.lots, D=P.days;
const days=(a,b)=>Math.round((new Date(b)-new Date(a))/864e5)+1;
const lotRows=L.map(l=>[l.name,n(l.kg),!l.arr?'In stock':md(l.arr),
  !l.arr?'(already dry)':`${md(l.arr)} – ${md(l.dry_end)}`,md(l.prep),md(l.ready),
  `${md(l.start)} – ${md(l.end)}`,(l.kg/560).toFixed(1)]);
const total=L.reduce((a,l)=>a+l.kg,0);
// appendix
const names=L.map(l=>l.name);
const app=D.map(d=>{const tot=names.reduce((a,k)=>a+d.stock[k],0);
  const s=Object.entries(d.sown).map(([k,v])=>`${k} ${n(v)}`).join(' / ')||'—';
  return [mdw(d.date,d.wd),...names.map(k=>d.stock[k]?n(d.stock[k]):'—'),n(tot),n(d.used),s]});
const shortSet=new Set(D.filter(d=>d.cap>0&&d.used<d.cap&&d.date>='2026-10-21'&&d.date<'2026-12-28').map(d=>mdw(d.date,d.wd)));
const sunSet=new Set(D.filter(d=>d.wd===6).map(d=>mdw(d.date,d.wd)));
const shortFill=(r)=>shortSet.has(r[0])?'FBE3E3':(sunSet.has(r[0])?'EDEDED':undefined);

const doc=new Document({
 styles:{default:{document:{run:{font:FONT,size:21}}},
  paragraphStyles:[
   {id:'Heading1',name:'Heading 1',basedOn:'Normal',next:'Normal',quickFormat:true,run:{size:28,bold:true,color:'1F4E79',font:FONT},paragraph:{spacing:{before:300,after:120},outlineLevel:0}},
   {id:'Heading2',name:'Heading 2',basedOn:'Normal',next:'Normal',quickFormat:true,run:{size:24,bold:true,color:'2E75B6',font:FONT},paragraph:{spacing:{before:200,after:100},outlineLevel:1}}]},
 numbering:{config:[{reference:'bul',levels:[{level:0,format:LevelFormat.BULLET,text:'•',alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:440,hanging:260}}}}]}]},
 sections:[{properties:{page:{margin:{top:1080,bottom:1080,left:1080,right:1080}}},children:[
  new Paragraph({children:[t('Jatropha Seed Storage and Sowing Plan',{size:36,bold:true,color:'1F4E79'})],spacing:{after:60}}),
  p([t('Period: 15 Oct – 28 Dec 2026   /   Updated: 16 Oct 2026',{color:'666666',size:19})],{border:{bottom:{style:BorderStyle.SINGLE,size:8,color:'1F4E79',space:4}},spacing:{after:200}}),

  h('1. Key points',HeadingLevel.HEADING_1),
  b([t('Total seed: '),t(`${n(total)} kg`,{bold:true}),t('. Weekly sowing capacity is 3,080 kg (560 kg × 5 weekdays + 280 kg on Saturday), so this is about 9.7 weeks of work.')]),
  b([t('Sowing period: '),t('Wed 21 Oct – Mon 28 Dec',{bold:true}),t('. The finish date is set by the arrival of Ghana 2 and Ghana 3, so starting on 21 Oct does not delay it.')]),
  b([t('Peak storage: '),t('18,720 kg (14–15 Nov)',{bold:true}),t('. India (1,020 kg left) and Ghana 1 (3,000 kg) are still in stock when Ghana 3 (13 Nov) and Ghana 2 (14 Nov) arrive.')]),
  b([t('Seed shortage: '),t('almost none',{bold:true}),t('. After sowing starts, the only shortfall is 60 kg on Sat 21 Nov (Ghana 1 runs out and Ghana 2 is ready from the next day). Sowing continues without a break from 21 Oct to 28 Dec.')]),
  b('Reserve storage space for at least 19 tonnes, including the drying area.'),

  h('2. Assumptions',HeadingLevel.HEADING_1),
  b('100 workers sow 560 kg per day in total across all lots. Saturday is a half day (280 kg); Sunday is off.'),
  b('Lots are used in this order: Ribáuè 2026 → Ribáuè 2025 → Ghana 0 → India → Ghana 1 → Ghana 2 → Ghana 3. If a lot runs out partway through a day, the rest of that day\'s 560 kg comes from the next lot.'),
  b('Each new arrival dries for 7 days, counting the arrival day (drying continues on Sundays). Day 8 is for preparation and soaking, and sowing can start on day 9. None of the preparation days falls on a Sunday.'),
  b('First sowing day is Wed 21 Oct. The lots already in stock (Ribáuè 2026, Ribáuè 2025, Ghana 0) are already dry and are prepared and soaked on Tue 20 Oct.'),
  b('Preparation and soaking is lead time for each lot; other lots can still be sown on that day.'),
  b('Storage is the amount left in the warehouse (including the drying area) after each day\'s sowing. Public holidays are not included (see section 5).'),

  h('3. Schedule by lot',HeadingLevel.HEADING_1),
  table(['Lot','kg','Arrival','Drying','Prep & soak','Ready','Sowing','Days at 560 kg'],lotRows,[1450,850,950,1300,1050,1050,1750,1200]),
  p([t('Note: India is ready from 24 Oct, Ghana 1 from 29 Oct and Ghana 3 from 21 Nov. To keep the set order, they are sown from 27 Oct, 16 Nov and 18 Dec.',{size:18,color:'555555'})],{spacing:{before:80,after:100}}),

  new Paragraph({children:[new PageBreak()]}),
  h('4. Storage over time',HeadingLevel.HEADING_1),
  new Paragraph({alignment:AlignmentType.CENTER,children:[new ImageRun({type:'png',data:fs.readFileSync(path.join(HERE,'stock_en.png')),transformation:{width:650,height:352}})]}),
  p([t('Figure 1. Storage by lot after each day\'s sowing. The dotted line marks the first sowing day (21 Oct). Light red bands are working days without enough seed ready. Nothing is sown on Sundays, so storage stays the same as on Saturday.',{size:18,color:'555555'})],{alignment:AlignmentType.CENTER}),

  h('4.1 Changes in storage',HeadingLevel.HEADING_2),
  table(['Period','Storage (kg)','What happens'],[
   ['15 Oct','2,820','Only the lots already in stock: Ribáuè 2026, Ribáuè 2025, Ghana 0'],
   ['16–20 Oct','11,220','India (8,400 kg) arrives and starts drying. No sowing yet; stock lots are prepared and soaked on 20 Oct'],
   ['21 Oct','13,660','Ghana 1 (3,000 kg) arrives. Sowing starts with Ribáuè 2026'],
   ['22–26 Oct','13,100 → 11,700','Ribáuè 2026 (finished 23 Oct) and Ribáuè 2025 (finished 26 Oct) are sown; Ghana 0 starts'],
   ['27 Oct','11,140','Ghana 0 finished; India starts'],
   ['28 Oct – 12 Nov','10,580 → 3,860','India is sown'],
   ['13 Nov','7,000','Ghana 3 (3,700 kg) arrives'],
   ['14–15 Nov','18,720','Ghana 2 (12,000 kg) arrives (peak)'],
   ['16–21 Nov','18,160 → 15,700','India finished 16 Nov; Ghana 1 sown and finished 21 Nov'],
   ['22 Nov','15,700','Sunday'],
   ['23 Nov – 18 Dec','15,140 → 3,660','Ghana 2 is sown'],
   ['18–28 Dec','3,660 → 0','Ghana 3 is sown; sowing complete on 28 Dec'],
  ],[1900,2000,5800]),

  h('4.2 Seed shortages',HeadingLevel.HEADING_2),
  table(['Date','Sown (kg)','Short (kg)','Reason'],[
   ['Sat 21 Nov','220','60','Ghana 1 runs out; Ghana 2 is ready from 22 Nov (sown from 23 Nov)'],
   ['Total','','60','Almost no impact'],
  ],[2000,1300,1300,5100],{fill:(r)=>r[0]==='Total'?'F2F2F2':undefined}),
  p([t('Note: 15–20 Oct is before sowing starts and is not counted as a shortage.',{size:18,color:'555555'})],{spacing:{before:80,after:100}}),

  h('5. Notes and suggestions',HeadingLevel.HEADING_1),
  b([t('Finish date is 28 Dec: ',{bold:true}),t('From 23 Nov, the 15,700 kg of Ghana 2 and Ghana 3 is sown without a break, so moving the start from 17 Oct to 21 Oct does not change the finish date. It also removes the idle working days of the earlier plan (23 Oct and 20–21 Nov).')]),
  b([t('Higher peak: ',{bold:true}),t('Because sowing starts later, India and Ghana 1 stay in stock until mid-November. The peak rises from 17,580 kg to 18,720 kg (+1,140 kg).')]),
  b([t('Before sowing starts: ',{bold:true}),t('16–20 Oct can be used for preparation, such as digging planting holes, preparing fields and briefing workers.')]),
  b([t('Holidays: ',{bold:true}),t('25 Dec (Family Day) is counted as a working day. If it is a day off, sowing finishes on Tue 29 Dec.')]),
  b([t('Drying space: ',{bold:true}),t('On 21–22 Oct, 11,400 kg (India and Ghana 1) is drying at the same time; on 14–19 Nov, 15,700 kg (Ghana 2 and Ghana 3). Size the drying area for these amounts.')]),
  b([t('How much to soak: ',{bold:true}),t('Soak only the next day\'s seed (560 kg on weekdays, 280 kg for Saturday) the day before. Because Sunday falls between Saturday and Monday, decide who soaks Monday\'s seed on Sunday, or soak it on Monday morning.')]),
  b([t('Lots stored for a long time: ',{bold:true}),t('India is kept for up to about 1 month after arrival (until 16 Nov), Ghana 1 for about 1 month (until 21 Nov) and Ghana 3 for about 1.5 months (until 28 Dec). Protect them from moisture and pests and inspect them regularly.')]),

  new Paragraph({children:[new PageBreak()]}),
  h('Appendix: daily storage and sowing (kg)',HeadingLevel.HEADING_1),
  p([t('Storage is after each day\'s sowing. Light red rows are days without enough seed; grey rows are Sundays (day off).',{size:18,color:'555555'})]),
  table(['Date','Rib 2026','Rib 2025','Ghana 0','India','Ghana 1','Ghana 2','Ghana 3','Total','Sown','Lots sown'],app,
    [1120,720,720,720,720,720,720,720,780,700,2040],{size:16,fill:shortFill}),
 ]}]});
Packer.toBuffer(doc).then(buf=>{fs.writeFileSync(OUT,buf);fs.copyFileSync(path.join(HERE,'stock_en.png'),path.join(HERE,'..','Seed_Storage_Chart_2026_EN.png'));console.log('wrote',OUT);});
