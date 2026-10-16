const fs=require('fs');
const path=require('path');
const HERE=__dirname;
const OUT=path.join(HERE,'..','種子保管・播種計画_2026.docx');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,ShadingType,AlignmentType,HeadingLevel,ImageRun,LevelFormat,BorderStyle,PageBreak}=require('docx');
const P=JSON.parse(fs.readFileSync(path.join(HERE,'plan.json')));
const WD='月火水木金土日';
const mdw=(s,w)=>md(s)+'('+WD[w]+')';
const md=s=>{if(!s||s==='None')return '—';const [y,m,d]=s.split('-');return `${+m}/${+d}`};
const n=v=>v.toLocaleString('en-US');
const FONT={ascii:'Arial',hAnsi:'Arial',eastAsia:'Yu Gothic'};
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
    children:[new Paragraph({alignment:hd?AlignmentType.CENTER:(/^[\d,.\-—+ kg日%/〜→]+$/.test(s)?AlignmentType.RIGHT:AlignmentType.LEFT),
      children:[t(s,{bold:!!hd,color:hd?'FFFFFF':undefined,size:opts.size||18})]})]});
  return new Table({width:{size:total,type:WidthType.DXA},columnWidths:widths,
    rows:[new TableRow({tableHeader:true,children:head.map((s,i)=>cell(s,widths[i],true))}),
      ...rows.map((r,ri)=>new TableRow({children:r.map((s,i)=>cell(String(s),widths[i],false,opts.fill?opts.fill(r,ri):undefined))}))]});
}
const L=P.lots, D=P.days;
const days=(a,b)=>Math.round((new Date(b)-new Date(a))/864e5)+1;
const lotRows=L.map(l=>[l.name,n(l.kg),!l.arr?'在庫あり':md(l.arr),
  !l.arr?'（乾燥済み）':`${md(l.arr)}〜${md(l.dry_end)}`,md(l.prep),md(l.ready),
  `${md(l.start)}〜${md(l.end)}`,(l.kg/560).toFixed(1)]);
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
  new Paragraph({children:[t('ジャトロファ種子 保管・播種計画',{size:36,bold:true,color:'1F4E79'})],spacing:{after:60}}),
  p([t('対象期間：2026年10月15日〜12月28日　／　更新日：2026年10月16日',{color:'666666',size:19})],{border:{bottom:{style:BorderStyle.SINGLE,size:8,color:'1F4E79',space:4}},spacing:{after:200}}),

  h('1. 要点',HeadingLevel.HEADING_1),
  b([t('総種子量 '),t(`${n(total)} kg`,{bold:true}),t('。1週間の播種能力は3,080 kg（平日560 kg×5日＋土曜280 kg）で、約9.7週分です。')]),
  b([t('播種期間：'),t('10/21（水）〜12/28（月）',{bold:true}),t('。完了日はGhana 2・3の到着日で決まるため、播種開始を10/21に遅らせても変わりません。')]),
  b([t('倉庫の最大保管量：'),t('18,720 kg（11/14〜11/15）',{bold:true}),t('。Indiaの播種中（残り1,020 kg）とGhana 1（3,000 kg）が残っているところに、Ghana 3（11/13）とGhana 2（11/14）が続けて到着するためです。')]),
  b([t('種子不足：'),t('ほぼなし',{bold:true}),t('。播種開始後に種子が足りなくなるのは11/21（土）の60 kgだけです（Ghana 1が終わり、Ghana 2は翌日から播種可能）。10/21〜12/28は途切れずに播種できます。')]),
  b('倉庫には、乾燥場所を含めて少なくとも19トン分の保管スペースを確保することを推奨します。'),

  h('2. 前提条件',HeadingLevel.HEADING_1),
  b('労働者100人で、1日の播種量は全ロット合計で560 kg。土曜日は半分の280 kg、日曜日は休み。'),
  b('ロットは指定の順番で使います：Ribáuè 2026 → Ribáuè 2025 → Ghana 0 → India → Ghana 1 → Ghana 2 → Ghana 3。前のロットがその日の途中で終わったら、残りの量は次のロットで播種します（合計は560 kgのまま）。'),
  b('到着日を含め7日間乾燥（日曜日も乾燥は続く）→ 8日目に準備・浸水 → 9日目から播種可能。今回は準備・浸水の日がどれも日曜日にあたりません。'),
  b('最初の播種は10/21（水）。既存のRibáuè 2026・2025・Ghana 0は乾燥済みとし、10/20（火）に準備・浸水します。'),
  b('準備・浸水はロット単位の準備期間として扱い、その間も他ロットの播種は続けられるものとします。'),
  b('保管量は「各日の播種後に倉庫（乾燥場所を含む）に残っている量」です。祝日は考慮していません（5章参照）。'),

  h('3. ロット別スケジュール',HeadingLevel.HEADING_1),
  table(['ロット','量 (kg)','到着日','乾燥','準備・浸水','播種可能日','播種期間','560kg換算日数'],lotRows,[1450,850,950,1300,1050,1050,1750,1200]),
  p([t('※ Indiaは10/24から、Ghana 1は10/29から、Ghana 3は11/21から播種可能ですが、指定の順番に従い、それぞれ10/27、11/16、12/18からの播種です。',{size:18,color:'555555'})],{spacing:{before:80,after:100}}),

  new Paragraph({children:[new PageBreak()]}),
  h('4. 倉庫保管量の推移',HeadingLevel.HEADING_1),
  new Paragraph({alignment:AlignmentType.CENTER,children:[new ImageRun({type:'png',data:fs.readFileSync(path.join(HERE,'stock.png')),transformation:{width:650,height:352}})]}),
  p([t('図1　種類別の保管量（各日の播種後）。点線は播種開始日（10/21）。薄い赤の帯は、作業日なのに播種できる種子が足りない日。日曜日は播種しないため、保管量が前日と同じになります。',{size:18,color:'555555'})],{alignment:AlignmentType.CENTER}),

  h('4.1 保管量の増減',HeadingLevel.HEADING_2),
  table(['期間','保管量 (kg)','動き'],[
   ['10/15','2,820','既存のRibáuè 2026・2025、Ghana 0のみ'],
   ['10/16〜10/20','11,220','India 8,400 kg 到着（乾燥開始）。播種前（10/20に既存分の準備・浸水）'],
   ['10/21','13,660','Ghana 1 3,000 kg 到着。播種開始（Ribáuè 2026）'],
   ['10/22〜10/26','13,100 → 11,700','Ribáuè 2026（10/23終了）、Ribáuè 2025（10/26終了）を播種。続けてGhana 0 を開始'],
   ['10/27','11,140','Ghana 0 が終了し、India を開始'],
   ['10/28〜11/12','10,580 → 3,860','India を播種'],
   ['11/13','7,000','Ghana 3 3,700 kg 到着'],
   ['11/14〜11/15','18,720','Ghana 2 12,000 kg 到着（最大）'],
   ['11/16〜11/21','18,160 → 15,700','India が11/16で終了。Ghana 1 を播種し11/21で終了'],
   ['11/22','15,700','日曜日'],
   ['11/23〜12/18','15,140 → 3,660','Ghana 2 を播種'],
   ['12/18〜12/28','3,660 → 0','Ghana 3 を播種し完了（12/28）'],
  ],[1900,2000,5800]),

  h('4.2 種子が足りない時期',HeadingLevel.HEADING_2),
  table(['日付','播種量 (kg)','不足分 (kg)','原因'],[
   ['11/21（土）','220','60','Ghana 1が終了。Ghana 2は11/22から播種可能（11/23から播種）'],
   ['合計','','60','ほぼ影響なし'],
  ],[2000,1300,1300,5100],{fill:(r)=>r[0]==='合計'?'F2F2F2':undefined}),
  p([t('※ 10/15〜10/20は播種開始前のため、不足には含めていません。',{size:18,color:'555555'})],{spacing:{before:80,after:100}}),

  h('5. 留意点・提案',HeadingLevel.HEADING_1),
  b([t('完了日は12/28：',{bold:true}),t('11/23以降はGhana 2・3の15,700 kgを続けて播種するため、開始日を10/17から10/21に遅らせても完了日は変わりません。その代わり、前回計算であった播種なしの作業日（10/23、11/20〜11/21）がなくなります。')]),
  b([t('最大保管量が増える：',{bold:true}),t('播種開始が遅れた分、IndiaとGhana 1の在庫が11月中旬まで残るため、最大保管量は17,580 kgから18,720 kgに増えます（+1,140 kg）。')]),
  b([t('播種前の期間：',{bold:true}),t('10/16〜10/20は、植穴掘り・圃場整備・労働者への説明などの準備に充てられます。')]),
  b([t('祝日：',{bold:true}),t('12/25（家族の日）は作業日として計算しています。休みにする場合は完了が12/29（火）になります。')]),
  b([t('乾燥スペース：',{bold:true}),t('10/21〜10/22はIndiaとGhana 1の計11,400 kg、11/14〜11/19はGhana 2・3の計15,700 kgが同時に乾燥中となります。乾燥場所の面積はこの量に合わせて確保してください。')]),
  b([t('浸水の量：',{bold:true}),t('浸水は翌日の播種量だけ（平日分560 kg、土曜分280 kg）を前日に行います。土曜日の播種後は日曜日を挟むため、月曜分の浸水は日曜日に誰が行うか、または月曜朝に行うかを決めておく必要があります。')]),
  b([t('長期保管ロット：',{bold:true}),t('Indiaは到着から最長約1か月（〜11/16）、Ghana 1は約1か月（〜11/21）、Ghana 3は約1.5か月（〜12/28）保管されます。湿気・害虫対策と定期点検を行ってください。')]),

  new Paragraph({children:[new PageBreak()]}),
  h('付録：日別の保管量と播種量（kg）',HeadingLevel.HEADING_1),
  p([t('保管量は各日の播種後の値。薄い赤の行は種子不足の日、灰色の行は日曜日（休み）。',{size:18,color:'555555'})]),
  table(['日付','Rib 2026','Rib 2025','Ghana 0','India','Ghana 1','Ghana 2','Ghana 3','合計','播種量','播種ロット'],app,
    [860,720,720,720,720,720,720,720,780,700,2300],{size:16,fill:shortFill}),
 ]}]});
Packer.toBuffer(doc).then(buf=>{fs.writeFileSync(OUT,buf);fs.copyFileSync(path.join(HERE,'stock.png'),path.join(HERE,'..','種子保管量グラフ_2026.png'));console.log('wrote',OUT);});
