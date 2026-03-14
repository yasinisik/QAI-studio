'use client'
import { useState, useRef, useEffect, useCallback } from 'react'
import { DEFAULT_TOOLS, CATEGORIES, PRESETS, PALETTE, buildPrompt } from '@/lib/tools'


// ─── MARKDOWN → HTML (basit) ─────────────────────────────────────────────────
function mdToHtmlSimple(md){
  let h = md
    .replace(/^### (.+)$/gm,'<h3>$1</h3>')
    .replace(/^## (.+)$/gm,'<h2>$1</h2>')
    .replace(/^# (.+)$/gm,'<h1>$1</h1>')
    .replace(/\*\*(.+?)\*\*/g,'<strong>$1</strong>')
    .replace(/\*(.+?)\*/g,'<em>$1</em>')
    .replace(/^> (.+)$/gm,'<blockquote>$1</blockquote>')
    .replace(/^- (.+)$/gm,'<li>$1</li>')
    .replace(/(<li>.*<\/li>\n?)+/g,'<ul>$&</ul>')
    .replace(/\n\n/g,'</p><p>')
  // Markdown tabloları HTML'e çevir
  h = h.replace(/\|(.+)\|/g, function(match){
    const cells = match.split('|').filter(c=>c.trim()&&!c.match(/^[-: ]+$/))
    if(!cells.length) return match
    return '<tr>' + cells.map(c=>'<td>'+c.trim()+'</td>').join('') + '</tr>'
  })
  h = h.replace(/(<tr>.*<\/tr>\n?)+/g,'<table>$&</table>')
  return '<p>' + h + '</p>'
}

// ─── DOCX EXPORT ──────────────────────────────────────────────────────────────
async function exportAsDocx(mdContent, title, fname){
  try {
    // docx kütüphanesini CDN'den yükle
    if(!window.docx){
      await new Promise((res,rej)=>{
        const sc = document.createElement('script')
        sc.src = 'https://cdnjs.cloudflare.com/ajax/libs/docx/7.8.2/docx.umd.min.js'
        sc.onload = res; sc.onerror = rej
        document.head.appendChild(sc)
      })
    }
    const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
            HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType } = window.docx

    const border = {style:BorderStyle.SINGLE,size:1,color:'AAAAAA'}
    const borders = {top:border,bottom:border,left:border,right:border}

    const children = []

    // Başlık
    children.push(new Paragraph({
      heading: HeadingLevel.TITLE,
      children:[new TextRun({text:title,bold:true,size:36,font:'Arial'})]
    }))

    // Markdown satırlarını parse et
    const lines = mdContent.split('\n')
    let tableRows = []
    let inTable = false

    for(let i=0; i<lines.length; i++){
      const line = lines[i]

      // Başlık
      if(line.startsWith('### '))  { children.push(new Paragraph({heading:HeadingLevel.HEADING_3,children:[new TextRun({text:line.slice(4),bold:true,font:'Arial',size:24})]})); continue }
      if(line.startsWith('## '))   { children.push(new Paragraph({heading:HeadingLevel.HEADING_2,children:[new TextRun({text:line.slice(3),bold:true,font:'Arial',size:28})]})); continue }
      if(line.startsWith('# '))    { children.push(new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun({text:line.slice(2),bold:true,font:'Arial',size:32})]})); continue }

      // Madde listesi
      if(line.startsWith('- ') || line.startsWith('* ')){
        children.push(new Paragraph({bullet:{level:0},children:[new TextRun({text:line.slice(2),font:'Arial',size:22})]}))
        continue
      }

      // Tablo satırı
      if(line.startsWith('|') && line.endsWith('|')){
        const cells = line.split('|').filter(c=>c.trim())
        if(cells.every(c=>/^[-: ]+$/.test(c))) continue // ayıraç satırı
        inTable = true
        const isHeader = tableRows.length===0
        const tr = new TableRow({
          children: cells.map(c=>new TableCell({
            borders,
            width:{size:Math.floor(9000/cells.length),type:WidthType.DXA},
            shading: isHeader?{fill:'1E3A8A',type:ShadingType.CLEAR}:undefined,
            margins:{top:80,bottom:80,left:120,right:120},
            children:[new Paragraph({children:[new TextRun({text:c.trim(),bold:isHeader,color:isHeader?'FFFFFF':'000000',font:'Arial',size:20})]})]
          }))
        })
        tableRows.push(tr)
        continue
      } else if(inTable && tableRows.length>0){
        children.push(new Table({width:{size:9000,type:WidthType.DXA},columnWidths:Array(tableRows[0].cells?.length||3).fill(Math.floor(9000/(tableRows[0].cells?.length||3))),rows:tableRows}))
        children.push(new Paragraph({children:[new TextRun({text:'',size:22})]}))
        tableRows = []; inTable = false
      }

      // Boş satır
      if(!line.trim()){
        children.push(new Paragraph({children:[new TextRun({text:'',size:22})]}))
        continue
      }

      // Bold text işle
      const parts = line.split(/(\*\*[^*]+\*\*)/g)
      const runs = parts.map(p=>{
        if(p.startsWith('**')&&p.endsWith('**')){
          return new TextRun({text:p.slice(2,-2),bold:true,font:'Arial',size:22})
        }
        return new TextRun({text:p,font:'Arial',size:22})
      })
      children.push(new Paragraph({children:runs}))
    }

    // Kalan tablo varsa ekle
    if(tableRows.length>0){
      children.push(new Table({width:{size:9000,type:WidthType.DXA},rows:tableRows}))
    }

    const doc = new Document({
      styles:{default:{document:{run:{font:'Arial',size:22}}}},
      sections:[{
        properties:{page:{size:{width:11906,height:16838},margin:{top:1200,right:1200,bottom:1200,left:1200}}},
        children
      }]
    })

    const blob = await Packer.toBlob(doc)
    const url = URL.createObjectURL(blob)
    const a = Object.assign(document.createElement('a'),{href:url,download:fname+'.docx'})
    a.click()
    URL.revokeObjectURL(url)
  } catch(e){
    alert('DOCX olusturulamadı: ' + e.message + '. Markdown veya HTML formatını deneyin.')
  }
}

// ─── STORAGE ──────────────────────────────────────────────────────────────────
const LS = {
  get: (k, d) => { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d } catch { return d } },
  set: (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)) } catch {} },
}

async function callAI(messages) {
  const res = await fetch('/api/ai', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({messages}) })
  const data = await res.json()
  if (data.error) throw new Error(data.error)
  return data.text
}

function exportTXT(results) {
  const txt = results.map(r=>`${r.tool.icon} ${r.tool.label}\n${'─'.repeat(60)}\n${r.content}`).join('\n\n')
  const a = Object.assign(document.createElement('a'),{href:URL.createObjectURL(new Blob([txt],{type:'text/plain;charset=utf-8'})),download:`qa-studio-${new Date().toISOString().slice(0,10)}.txt`})
  a.click()
}

// ─── DESIGN SYSTEM ────────────────────────────────────────────────────────────
const C = {
  // Backgrounds
  bg0: '#05080f',
  bg1: '#090d18',
  bg2: '#0d1220',
  bg3: '#111827',
  // Borders
  b0: 'rgba(255,255,255,0.05)',
  b1: 'rgba(255,255,255,0.09)',
  b2: 'rgba(99,179,237,0.22)',
  // Text
  t0: '#f0f4ff',
  t1: '#c8d3e8',
  t2: '#7a8fa8',
  t3: '#3d5070',
  // Accents
  blue:   '#3b9eed',
  purple: '#8b5cf6',
  green:  '#10b981',
  red:    '#ef4444',
  orange: '#f97316',
  yellow: '#f59e0b',
  pink:   '#ec4899',
  teal:   '#14b8a6',
  gold:   '#fbbf24',
  cyan:   '#22d3ee',
  lime:   '#a3e635',
  rose:   '#fb7185',
  indigo: '#6366f1',
  cyan:   '#06b6d4',
  lime:   '#84cc16',
  amber:  '#f59e0b',
  rose:   '#f43f5e',
  indigo: '#6366f1',
  sky:    '#38bdf8',
  emerald:'#34d399',
}

const SEV = {
  Highest: { fg:'#ff4040', bg:'rgba(255,64,64,0.1)',  ring:'rgba(255,64,64,0.3)',  icon:'⬆️' },
  High:    { fg:'#f87171', bg:'rgba(248,113,113,0.1)', ring:'rgba(248,113,113,0.3)', icon:'🔴' },
  Medium:  { fg:'#fb923c', bg:'rgba(251,146,60,0.1)',  ring:'rgba(251,146,60,0.3)',  icon:'🟠' },
  Low:     { fg:'#fbbf24', bg:'rgba(251,191,36,0.1)',  ring:'rgba(251,191,36,0.3)',  icon:'🟡' },
  Lowest:  { fg:'#34d399', bg:'rgba(52,211,153,0.1)',  ring:'rgba(52,211,153,0.3)',  icon:'⬇️' },
}

const STA = {
  Open:          { fg:'#3b9eed', bg:'rgba(59,158,237,0.1)'  },
  'In Progress': { fg:'#8b5cf6', bg:'rgba(139,92,246,0.1)'  },
  Resolved:      { fg:'#10b981', bg:'rgba(16,185,129,0.1)'  },
  Closed:        { fg:'#7a8fa8', bg:'rgba(122,143,168,0.1)' },
}

// ─── ATOMS ────────────────────────────────────────────────────────────────────
const Spin = ({c=C.blue,s=14})=><span style={{display:'inline-block',width:s,height:s,border:'2px solid rgba(255,255,255,0.07)',borderTopColor:c,borderRadius:'50%',animation:'qa-spin .65s linear infinite',flexShrink:0,verticalAlign:'middle'}}/>

function Pill({children, color=C.blue, onClick, active}){
  const [h,sH]=useState(false)
  return <span onClick={onClick} style={{display:'inline-flex',alignItems:'center',gap:4,padding:'3px 10px',borderRadius:99,fontSize:11,fontWeight:600,cursor:onClick?'pointer':'default',background:active||h?`${color}1a`:C.bg3,border:`1px solid ${active||h?color+'55':C.b0}`,color:active||h?color:C.t2,transition:'all .15s'}}>{children}</span>
}

function Tag({children, color=C.blue}){
  return <span style={{display:'inline-flex',alignItems:'center',padding:'2px 8px',borderRadius:5,fontSize:11,fontWeight:700,background:`${color}18`,border:`1px solid ${color}35`,color}}>{children}</span>
}

// Primary action button — blue gradient, subtle glow
function BtnPrimary({children, onClick, disabled, loading, icon, size='md'}){
  const [h,sH]=useState(false)
  const p = size==='sm'?'6px 14px':size==='lg'?'12px 22px':'9px 18px'
  const fs = size==='sm'?12:size==='lg'?14:13
  return(
    <button onClick={onClick} disabled={disabled||loading} onMouseEnter={()=>sH(true)} onMouseLeave={()=>sH(false)}
      style={{display:'inline-flex',alignItems:'center',gap:6,padding:p,borderRadius:8,border:'1px solid rgba(59,158,237,0.3)',background:disabled||loading?C.bg3:h?'rgba(59,158,237,0.22)':'rgba(59,158,237,0.14)',color:disabled?C.t3:C.blue,cursor:disabled||loading?'not-allowed':'pointer',fontFamily:'inherit',fontWeight:600,fontSize:fs,boxShadow:disabled||loading?'none':'0 0 0 0px rgba(59,158,237,0)',transition:'all .15s',whiteSpace:'nowrap'}}>
      {loading?<Spin c={C.blue} s={fs}/>:icon&&<span style={{fontSize:fs}}>{icon}</span>}
      {children}
    </button>
  )
}

// Secondary/ghost button
function BtnGhost({children, onClick, disabled, icon, size='md', color}){
  const [h,sH]=useState(false)
  const fg = color||C.t2
  const p = size==='sm'?'5px 12px':size==='lg'?'11px 20px':'8px 16px'
  const fs = size==='sm'?12:size==='lg'?14:13
  return(
    <button onClick={onClick} disabled={disabled} onMouseEnter={()=>sH(true)} onMouseLeave={()=>sH(false)}
      style={{display:'inline-flex',alignItems:'center',gap:6,padding:p,borderRadius:8,border:`1px solid ${h?C.b1:C.b0}`,background:h?C.bg3:'transparent',color:disabled?C.t3:h?C.t1:fg,cursor:disabled?'not-allowed':'pointer',fontFamily:'inherit',fontWeight:600,fontSize:fs,transition:'all .15s',whiteSpace:'nowrap'}}>
      {icon&&<span style={{fontSize:fs}}>{icon}</span>}
      {children}
    </button>
  )
}

// Danger button
function BtnDanger({children, onClick, size='md', icon}){
  const [h,sH]=useState(false)
  const p = size==='sm'?'5px 11px':'8px 16px'
  return(
    <button onClick={onClick} onMouseEnter={()=>sH(true)} onMouseLeave={()=>sH(false)}
      style={{display:'inline-flex',alignItems:'center',gap:5,padding:p,borderRadius:7,border:`1px solid ${h?'rgba(239,68,68,.4)':'rgba(239,68,68,.2)'}`,background:h?'rgba(239,68,68,.15)':'rgba(239,68,68,.07)',color:'#f87171',cursor:'pointer',fontFamily:'inherit',fontWeight:600,fontSize:12,transition:'all .15s'}}>
      {icon&&<span>{icon}</span>}{children}
    </button>
  )
}

// Success button
function BtnSuccess({children, onClick, size='md', icon}){
  const [h,sH]=useState(false)
  const p = size==='sm'?'5px 12px':'9px 18px'
  const fs = size==='sm'?12:13
  return(
    <button onClick={onClick} onMouseEnter={()=>sH(true)} onMouseLeave={()=>sH(false)}
      style={{display:'inline-flex',alignItems:'center',gap:6,padding:p,borderRadius:8,border:`1px solid ${h?'rgba(16,185,129,.4)':'rgba(16,185,129,.2)'}`,background:h?'rgba(16,185,129,.15)':'rgba(16,185,129,.08)',color:C.green,cursor:'pointer',fontFamily:'inherit',fontWeight:600,fontSize:fs,transition:'all .15s',whiteSpace:'nowrap'}}>
      {icon&&<span style={{fontSize:fs}}>{icon}</span>}{children}
    </button>
  )
}

// Purple/accent button
function BtnAccent({children, onClick, disabled, loading, icon, size='md'}){
  const [h,sH]=useState(false)
  const p = size==='sm'?'6px 13px':'9px 18px'
  const fs = size==='sm'?12:13
  return(
    <button onClick={onClick} disabled={disabled||loading} onMouseEnter={()=>sH(true)} onMouseLeave={()=>sH(false)}
      style={{display:'inline-flex',alignItems:'center',gap:6,padding:p,borderRadius:8,border:`1px solid ${disabled?C.b0:'rgba(139,92,246,.3)'}`,background:disabled?C.bg3:h?'rgba(139,92,246,.2)':'rgba(139,92,246,.1)',color:disabled?C.t3:C.purple,cursor:disabled||loading?'not-allowed':'pointer',fontFamily:'inherit',fontWeight:600,fontSize:fs,transition:'all .15s',whiteSpace:'nowrap'}}>
      {loading?<Spin c={C.purple} s={fs}/>:icon&&<span style={{fontSize:fs}}>{icon}</span>}
      {children}
    </button>
  )
}

function Field({label, children}){
  return(
    <div>
      <div style={{fontSize:11,fontWeight:600,color:C.t3,textTransform:'uppercase',letterSpacing:.8,marginBottom:6}}>{label}</div>
      {children}
    </div>
  )
}

function TextInput({value, onChange, onKeyDown, placeholder, sx={}}){
  const [f,sF]=useState(false)
  return <input value={value} onChange={onChange} onKeyDown={onKeyDown} placeholder={placeholder}
    onFocus={()=>sF(true)} onBlur={()=>sF(false)}
    style={{width:'100%',background:'rgba(255,255,255,0.03)',border:`1px solid ${f?C.b2:C.b1}`,borderRadius:8,padding:'9px 13px',color:C.t0,fontSize:13,fontFamily:'inherit',outline:'none',transition:'border .15s',...sx}}/>
}

function TextArea({value, onChange, onKeyDown, placeholder, rows=4, sx={}}){
  const [f,sF]=useState(false)
  return <textarea value={value} onChange={onChange} onKeyDown={onKeyDown} placeholder={placeholder} rows={rows}
    onFocus={()=>sF(true)} onBlur={()=>sF(false)}
    style={{width:'100%',background:'rgba(255,255,255,0.03)',border:`1px solid ${f?C.b2:C.b1}`,borderRadius:8,padding:'10px 13px',color:C.t0,fontSize:13,fontFamily:'inherit',outline:'none',resize:'vertical',lineHeight:1.7,transition:'border .15s',...sx}}/>
}

function Sel({value, onChange, children, sx={}}){
  return <select value={value} onChange={onChange}
    style={{width:'100%',background:C.bg3,border:`1px solid ${C.b1}`,borderRadius:8,padding:'9px 12px',color:C.t0,fontSize:13,fontFamily:'inherit',outline:'none',...sx}}>
    {children}
  </select>
}

function Divider({sx={}}){
  return <div style={{height:1,background:C.b0,...sx}}/>
}




// ─── BENI OKU DOKÜMANI ────────────────────────────────────────────────────────
const TOOL_SETUP_INFO = {
  playwright: {
    title: 'Playwright Kurulum Rehberi',
    steps: [
      { h: 'Node.js Kurulumu', items: ['nodejs.org adresine git', 'LTS surumunu indir ve kur', 'CMD ac → node -v yazdiginda versiyon gorunmeli'] },
      { h: 'Playwright Kurulumu', items: ['CMD de proje klasorune gir', 'npm init -y komutunu calistir', 'npm install @playwright/test komutunu calistir', 'npx playwright install komutunu calistir (tarayicilari indirir)'] },
      { h: 'Testi Calistir', items: ['npx playwright test dosya_adi.spec.ts', 'Raporlar icin: npx playwright show-report'] },
    ],
    note: 'Playwright, Chrome, Firefox ve Safari otomatik olarak test edebilir.'
  },
  selenium: {
    title: 'Selenium Python Kurulum Rehberi',
    steps: [
      { h: 'Python Kurulumu', items: ['python.org adresine git', 'Python 3.10+ surumunu indir ve kur', 'Add to PATH secenegini isaretله', 'CMD ac → python --version yazdiginda versiyon gorunmeli'] },
      { h: 'Selenium Kurulumu', items: ['pip install selenium', 'pip install webdriver-manager'] },
      { h: 'Testi Calistir', items: ['python dosya_adi.py', 'Hata alirsan: pip install --upgrade selenium'] },
    ],
    note: 'Chrome WebDriver otomatik indirilir — webdriver-manager sayesinde manuel kurulum gerekmez.'
  },
  appium: {
    title: 'Appium Mobil Test Kurulum Rehberi',
    steps: [
      { h: 'Gereksinimler', items: ['Java JDK 11+ kur (adoptium.net)', 'Android Studio kur (Android icin)', 'Xcode kur (iOS icin, sadece Mac sistemlerde)', 'Node.js kur (nodejs.org)'] },
      { h: 'Appium Kurulumu', items: ['npm install -g appium', 'appium driver install uiautomator2 (Android)', 'appium driver install xcuitest (iOS)'] },
      { h: 'Appium Baslatma ve Test', items: ['CMD: appium yaz (sunucu baslar)', 'Yeni CMD penceresi: python dosya_adi.py', 'Emulator veya gercek cihaz bagli olmali'] },
    ],
    note: 'Android testi icin AVD Manager ile emulator olustur veya USB ile gercek cihaz bagla.'
  },
  jmeter: {
    title: 'JMeter / k6 Performans Testi Kurulum Rehberi',
    steps: [
      { h: 'Java Kurulumu (JMeter icin)', items: ['adoptium.net adresine git', 'Temurin 21 LTS indir ve kur', 'CMD → java -version ile dogrula'] },
      { h: 'JMeter Indir', items: ['jmeter.apache.org/download_jmeter.cgi', 'Binary ZIP dosyasini indir', "ZIP'i ac — orn: C:\\jmeter\\", 'C:\\jmeter\\bin\\ klasorune dosyayi koy'] },
      { h: '.js Dosyasini .jmx Olarak Kaydet', items: ['Indirilen dosyanin adindaki .js uzantisini .jmx olarak degistir', 'Orn: test.js → test.jmx'] },
      { h: 'Testi Calistir (CMD)', items: ['cd C:\\jmeter\\bin', 'jmeter -n -t test.jmx -l sonuc.jtl -e -o rapor', 'Raporu ac: rapor\\index.html'] },
    ],
    note: '1000 kullanici gibi buyuk yukleri her zaman komut satirinda (-n) calistir. GUI modu sadece plan olusturmak icindir.'
  },
  batch: {
    title: 'Batch / Shell Script Calistirma Rehberi',
    steps: [
      { h: 'Windows .bat Dosyasi', items: ['Indirilen dosyayi .bat olarak kaydet', 'Sag tikla → Yonetici olarak calistir', 'Veya CMD: dosya_adi.bat'] },
      { h: 'Linux/Mac .sh Dosyasi', items: ['Terminal ac', 'chmod +x dosya_adi.sh', './dosya_adi.sh'] },
      { h: 'Hata Alirsan', items: ['Windows: PowerShell yerine CMD kullan', 'Yetki hatasi: Yonetici olarak calistir', 'Linux: bash dosya_adi.sh ile dene'] },
    ],
    note: 'Script icindeki yollari kendi sisteminize gore duzenleyin.'
  },
  html_report: {
    title: 'HTML Test Raporu Goruntuleme Rehberi',
    steps: [
      { h: 'Dosyayi Kaydet', items: ['Indirilen .html dosyasini bilgisayara kaydet'] },
      { h: 'Tarayicide Ac', items: ['Dosyaya cift tikla — dogrudan tarayicide acilir', 'Veya Chrome: Ctrl+O ile dosyayi sec'] },
      { h: 'Paylasma', items: ['Dosyayi mail ile gonder', 'Confluence veya SharePoint e yukle', 'GitHub Pages ile yayinla'] },
    ],
    note: 'HTML raporu tek dosyadir — internet baglantisi gerekmez, her tarayicide calisir.'
  },
}
function downloadBeniOku(tool, content) {
  const info = TOOL_SETUP_INFO[tool.toolType]
  const toolLabel = tool.label || 'Arac'
  const date = new Date().toLocaleDateString('tr-TR')

  const css = [
    '*{box-sizing:border-box;margin:0;padding:0}',
    'body{font-family:Segoe UI,system-ui,sans-serif;background:#05080f;color:#c8d8e8;padding:40px;max-width:860px;margin:0 auto}',
    'h1{font-size:28px;font-weight:900;color:#3b9eed;margin-bottom:6px}',
    '.sub{color:#7a8fa8;font-size:14px;margin-bottom:32px;padding-bottom:16px;border-bottom:1px solid #1a2a4a}',
    'h2{font-size:16px;font-weight:700;color:#7ec8f0;margin:28px 0 12px;padding:8px 14px;background:#0d1a30;border-left:3px solid #3b9eed;border-radius:0 6px 6px 0}',
    'ol{padding-left:24px;margin-bottom:16px}',
    'li{margin-bottom:8px;line-height:1.6;font-size:14px}',
    '.note{background:#1a1200;border:1px solid #f59e0b44;border-radius:8px;padding:12px 16px;margin:24px 0;color:#f59e0b;font-size:13px}',
    '.note::before{content:"\u26A0 ";font-weight:700}',
    '.section{background:#090d18;border:1px solid #1a2a4a;border-radius:10px;padding:20px 24px;margin-bottom:16px}',
    '.code-box{background:#0a0f1e;border:1px solid #1a3060;border-radius:7px;padding:14px 18px;margin:16px 0;font-family:Consolas,monospace;font-size:13px;color:#a8f0c8;white-space:pre-wrap;word-break:break-all}',
    '.footer{margin-top:48px;padding-top:16px;border-top:1px solid #1a2a4a;color:#3d5070;font-size:12px;text-align:center}',
    '.badge{display:inline-block;padding:3px 10px;border-radius:5px;font-size:11px;font-weight:700;margin-bottom:8px}'
  ].join(' ')

  let html = '<!DOCTYPE html><html lang="tr"><head><meta charset="UTF-8"><title>Beni Oku - ' + toolLabel + '</title><style>' + css + '</style></head><body>'
  html += '<div class="badge" style="background:#3b9eed22;color:#3b9eed;border:1px solid #3b9eed44">QAI Studio</div>'
  html += '<h1>' + (info ? info.title : 'Beni Oku - ' + toolLabel) + '</h1>'
  html += '<div class="sub">Arac: <strong style="color:#c8d8e8">' + toolLabel + '</strong> &nbsp;&middot;&nbsp; ' + date + '</div>'

  if (info) {
    info.steps.forEach(function(step, i) {
      const items = step.items.map(function(item){ return '<li>' + item + '</li>' }).join('')
      html += '<div class="section"><h2>' + (i+1) + '. ' + step.h + '</h2><ol>' + items + '</ol></div>'
    })
    if (info.note) html += '<div class="note">' + info.note + '</div>'
  }

  const safeContent = content.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').slice(0, 3000)
  const more = content.length > 3000 ? '\n... (devami icin Kopyala butonunu kullan)' : ''
  html += '<div class="section"><h2>Uretilen Cikti</h2><div class="code-box">' + safeContent + more + '</div></div>'

  if (!info) {
    html += '<div class="section"><h2>Arac Sonucu</h2><div class="code-box">' + safeContent + '</div></div>'
  }

  html += '<div class="footer">QAI Studio · AI Destekli Test Muhendisligi</div></body></html>'

  const fname = toolLabel.replace(/\s+/g,'_').toLowerCase() + '_sonuc.html'
  downloadFile(html, fname, 'text/html')
}

// ─── SPLASH SCREEN ────────────────────────────────────────────────────────────
function SplashScreen({exiting, onEnter}){
  const [loaded, setLoaded] = useState(false)
  useEffect(()=>{ setTimeout(()=>setLoaded(true), 80) },[])

  return(
    <div style={{
      position:'fixed', inset:0, background:'#04060e',
      display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center',
      overflow:'hidden', cursor:'pointer',
      animation: exiting ? 'splashZoomOut 1.35s cubic-bezier(.22,1,.36,1) forwards' : 'none',
    }} onClick={onEnter}>

      {/* Animated background grid */}
      <div style={{position:'absolute',inset:0,backgroundImage:`linear-gradient(rgba(59,158,237,0.04) 1px,transparent 1px),linear-gradient(90deg,rgba(59,158,237,0.04) 1px,transparent 1px)`,backgroundSize:'48px 48px',animation:'gridDrift 20s linear infinite'}}/>

      {/* Radial glow center */}
      <div style={{position:'absolute',top:'50%',left:'50%',transform:'translate(-50%,-50%)',width:600,height:600,background:'radial-gradient(circle,rgba(59,158,237,0.12) 0%,transparent 70%)',pointerEvents:'none'}}/>

      {/* Floating particles */}
      {[...Array(18)].map((_,i)=>(
        <div key={i} style={{
          position:'absolute',
          width: 2+Math.random()*3,
          height: 2+Math.random()*3,
          borderRadius:'50%',
          background:`rgba(59,158,237,${0.2+Math.random()*0.4})`,
          left:`${5+Math.random()*90}%`,
          top:`${5+Math.random()*90}%`,
          animation:`float${i%3} ${4+Math.random()*6}s ease-in-out infinite`,
          animationDelay:`${Math.random()*4}s`,
        }}/>
      ))}

      {/* Main content */}
      <div style={{
        position:'relative', textAlign:'center', zIndex:2,
        opacity: loaded ? 1 : 0,
        transform: loaded ? 'translateY(0)' : 'translateY(24px)',
        transition: 'opacity .7s ease, transform .7s ease',
      }}>
        {/* Icon */}
        <div style={{
          fontSize:72, lineHeight:1, marginBottom:24,
          filter:'drop-shadow(0 0 32px rgba(59,158,237,0.6))',
          animation:'iconPulse 3s ease-in-out infinite',
        }}>🧪</div>

        {/* Title */}
        <div style={{
          fontSize:56, fontWeight:900, letterSpacing:-2, lineHeight:1,
          color:'#f0f4ff', marginBottom:10,
          fontFamily:"'Segoe UI',system-ui,sans-serif",
          textShadow:'0 0 60px rgba(59,158,237,0.3)',
        }}>
          <span style={{background:'linear-gradient(135deg,#fbbf24,#f59e0b)',WebkitBackgroundClip:'text',WebkitTextFillColor:'transparent'}}>QAI</span>{' '}Studio
        </div>

        {/* Tagline */}
        <div style={{
          fontSize:16, color:'rgba(200,211,232,0.6)', fontWeight:400, letterSpacing:.5,
          marginBottom:48, fontFamily:"'Segoe UI',system-ui,sans-serif",
        }}>
          AI Destekli Test Mühendisliği Platformu
        </div>

        {/* CTA Button */}
        <div style={{
          display:'inline-flex', alignItems:'center', gap:10,
          padding:'14px 32px', borderRadius:12,
          background:'rgba(59,158,237,0.14)',
          border:'1px solid rgba(59,158,237,0.35)',
          color:'#3b9eed', fontSize:15, fontWeight:700,
          fontFamily:"'Segoe UI',system-ui,sans-serif",
          letterSpacing:.3,
          boxShadow:'0 0 30px rgba(59,158,237,0.15)',
          animation:'ctaPulse 2.5s ease-in-out infinite',
        }}>
          <span style={{fontSize:18}}>▶</span>
          Uygulamayı Aç
        </div>

        {/* Hint */}
        <div style={{marginTop:20,fontSize:12,color:'rgba(120,143,168,0.4)',fontFamily:"'Segoe UI',system-ui,sans-serif"}}>
          Herhangi Bir Yere Tıkla
        </div>
      </div>

      {/* Bottom version */}
      <div style={{position:'absolute',bottom:24,fontSize:11,color:'rgba(61,80,112,0.7)',fontFamily:"'Segoe UI',system-ui,sans-serif",letterSpacing:.5}}>
        QAI Studio v1.0  ·  AI Powered
      </div>

      <style>{`
        @keyframes gridDrift { from{backgroundPosition:0 0} to{backgroundPosition:48px 48px} }
        @keyframes float0 { 0%,100%{transform:translateY(0) translateX(0)} 50%{transform:translateY(-18px) translateX(8px)} }
        @keyframes float1 { 0%,100%{transform:translateY(0) translateX(0)} 50%{transform:translateY(12px) translateX(-10px)} }
        @keyframes float2 { 0%,100%{transform:translateY(0) translateX(0)} 50%{transform:translateY(-10px) translateX(-6px)} }
        @keyframes iconPulse { 0%,100%{filter:drop-shadow(0 0 28px rgba(59,158,237,0.5))} 50%{filter:drop-shadow(0 0 48px rgba(59,158,237,0.85))} }
        @keyframes ctaPulse { 0%,100%{box-shadow:0 0 20px rgba(59,158,237,0.15)} 50%{box-shadow:0 0 36px rgba(59,158,237,0.32)} }
        @keyframes splashZoomOut {
          0%   { transform:scale(1) translateZ(0);    opacity:1; filter:blur(0px) }
          15%  { transform:scale(1.04) translateZ(0); opacity:1; filter:blur(0px) }
          50%  { transform:scale(2.0) translateZ(0);  opacity:0.7; filter:blur(4px) }
          100% { transform:scale(5.0) translateZ(0);  opacity:0; filter:blur(24px) }
        }
      `}</style>
    </div>
  )
}



// ─── SIK KULLANILAN PANEL ─────────────────────────────────────────────────────
function SikKullanilanPanel({topIds, allTools, selectedTools, toggle}){
  const [open,setOpen] = useState(true)
  return(
    <div style={{padding:'0 10px 6px'}}>
      <div onClick={()=>setOpen(v=>!v)} style={{display:'flex',alignItems:'center',gap:5,marginBottom:open?6:0,cursor:'pointer',userSelect:'none',padding:'3px 0'}}>
        <span style={{fontSize:9,fontWeight:700,color:C.t3,textTransform:'uppercase',letterSpacing:1,flex:1}}>⭐ Sık Kullanılan</span>
        <span style={{fontSize:9,color:C.t3,transition:'transform .2s',display:'inline-block',transform:open?'rotate(90deg)':'rotate(0deg)'}}>▶</span>
      </div>
      {open&&(
        <div style={{display:'flex',flexWrap:'wrap',gap:4}}>
          {topIds.map(id=>{const t=allTools.find(x=>x.id===id);if(!t)return null;const cl=PALETTE[t.color]||PALETTE.blue;return <Pill key={id} color={cl.text} onClick={()=>toggle(id)} active={selectedTools.includes(id)}>{t.icon} {t.label}</Pill>})}
        </div>
      )}
    </div>
  )
}


// ─── DASHBOARD EXPORT ─────────────────────────────────────────────────────────
function DashExport({bugs}){
  const [fmt,setFmt] = useState('csv')

  const doExport = () => {
    const date = new Date().toISOString().slice(0,10)
    const fname = 'qa-bugs-' + date

    if(fmt==='csv'){
      const head = 'ID,Baslik,Severity,Durum,Versiyon,Tarih,Aciklama'
      const rows = bugs.map((b,i)=>[i+1, '"'+(b.title||'').replace(/"/g,'""')+'"', b.severity, b.status, b.version, b.date, '"'+(b.desc||'').replace(/"/g,'""')+'"'].join(','))
      downloadFile([head,...rows].join(NL), fname+'.csv', 'text/csv')
    }
    else if(fmt==='json'){
      downloadFile(JSON.stringify(bugs, null, 2), fname+'.json', 'application/json')
    }
    else if(fmt==='txt'){
      const lines = bugs.map((b,i)=>[(i+1)+'. ['+b.severity+'] '+(b.title||''), '   Durum: '+b.status+' | Versiyon: '+(b.version||'')+' | Tarih: '+(b.date||''), b.desc?'   '+b.desc:''].filter(Boolean).join(NL))
      downloadFile(lines.join(NL+NL), fname+'.txt', 'text/plain')
    }
    else if(fmt==='md'){
      const head = '| # | Başlık | Severity | Durum | Versiyon | Tarih |' + NL + '|---|--------|----------|-------|----------|-------|'
      const rows = bugs.map((b,i)=>'| '+(i+1)+' | '+(b.title||'')+' | '+b.severity+' | '+b.status+' | '+(b.version||'')+' | '+(b.date||'')+' |')
      downloadFile(['# Bug Raporu — '+new Date().toLocaleDateString('tr-TR'), '', head, ...rows].join(NL), fname+'.md', 'text/markdown')
    }
    else if(fmt==='html'){
      const sevColor = {Highest:'#ff4040',High:'#f87171',Medium:'#fb923c',Low:'#fbbf24',Lowest:'#34d399'}
      const staColor = {Open:'#3b9eed','In Progress':'#8b5cf6',Resolved:'#10b981',Closed:'#7a8fa8'}
      const rows = bugs.map((b,i)=>'<tr><td>'+(i+1)+'</td><td>'+(b.title||'')+'</td><td style="color:'+(sevColor[b.severity]||'#ccc')+'">'+b.severity+'</td><td style="color:'+(staColor[b.status]||'#ccc')+'">'+b.status+'</td><td>'+(b.version||'')+'</td><td>'+(b.date||'')+'</td><td>'+(b.desc||'')+'</td></tr>').join('')
      const html = '<!DOCTYPE html><html lang="tr"><head><meta charset="UTF-8"><title>QAI Bug Raporu</title><style>body{font-family:Segoe UI,sans-serif;background:#05080f;color:#c8d8e8;padding:32px}h1{color:#3b9eed;margin-bottom:4px}p{color:#7a8fa8;font-size:13px;margin-bottom:24px}table{width:100%;border-collapse:collapse}th{background:#0d1220;color:#7ec8f0;font-size:11px;text-transform:uppercase;padding:10px 14px;text-align:left;border-bottom:1px solid #1a2a4a}td{padding:9px 14px;border-bottom:1px solid #0d1220;font-size:13px}tr:hover td{background:rgba(255,255,255,0.02)}</style></head><body><h1>🐛 Bug Raporu</h1><p>QAI Studio · '+new Date().toLocaleDateString('tr-TR')+' · '+bugs.length+' bug</p><table><thead><tr><th>#</th><th>Başlık</th><th>Severity</th><th>Durum</th><th>Versiyon</th><th>Tarih</th><th>Açıklama</th></tr></thead><tbody>'+rows+'</tbody></table></body></html>'
      downloadFile(html, fname+'.html', 'text/html')
    }
  }

  return(
    <div style={{display:'flex',alignItems:'center',gap:4}}>
      <select value={fmt} onChange={e=>setFmt(e.target.value)}
        style={{fontSize:11,padding:'4px 7px',borderRadius:5,border:`1px solid ${C.b1}`,background:C.bg3,color:C.t2,fontFamily:'inherit',cursor:'pointer'}}>
        <option value="csv">📊 CSV</option>
        <option value="json">🔧 JSON</option>
        <option value="html">🌐 HTML</option>
        <option value="md">📝 Markdown</option>
        <option value="txt">📄 TXT</option>
      </select>
      <BtnSuccess size="sm" icon="📥" onClick={doExport}>İndir</BtnSuccess>
    </div>
  )
}

// ─── PRESET PANEL ─────────────────────────────────────────────────────────────
function PresetPanel({presets, allTools, selectedTools, setST, savePresets}){
  const [editMode, setEditMode] = useState(false)
  const [editIdx, setEditIdx]   = useState(null)  // düzenlenen preset index
  const [newLabel, setNewLabel] = useState('')
  const [newTools, setNewTools] = useState([])    // seçilen tool id'leri
  const [addMode, setAddMode]   = useState(false)

  const startEdit = (i) => {
    setEditIdx(i)
    setNewLabel(presets[i].label)
    setNewTools([...presets[i].tools])
    setAddMode(false)
  }
  const startAdd = () => {
    setEditIdx(null)
    setNewLabel('')
    setNewTools([...selectedTools])
    setAddMode(true)
  }
  const cancelForm = () => { setEditIdx(null); setAddMode(false); setNewLabel(''); setNewTools([]) }

  const toggleTool = (id) => setNewTools(t => t.includes(id) ? t.filter(x=>x!==id) : [...t,id])

  const saveForm = () => {
    if(!newLabel.trim()||!newTools.length) return
    let updated
    if(addMode){
      updated = [...presets, {label: newLabel.trim(), tools: newTools}]
    } else {
      updated = presets.map((p,i)=>i===editIdx?{...p, label:newLabel.trim(), tools:newTools}:p)
    }
    savePresets(updated)
    cancelForm()
  }

  const deletePreset = (i) => {
    savePresets(presets.filter((_,idx)=>idx!==i))
    if(editIdx===i) cancelForm()
  }

  const isFormOpen = addMode || editIdx!==null

  return(
    <div style={{marginBottom:14,background:'rgba(255,255,255,0.02)',borderRadius:8,border:`1px solid ${C.b0}`}}>
      {/* Başlık */}
      <div style={{display:'flex',alignItems:'center',gap:6,padding:'8px 12px',borderBottom:editMode||isFormOpen?`1px solid ${C.b0}`:'none'}}>
        <span style={{fontSize:10,fontWeight:700,color:C.t3,textTransform:'uppercase',letterSpacing:.8,flex:1}}>⚡ Hazır Paketler — hızlı başla</span>
        {!isFormOpen&&(
          <>
            <button onClick={()=>{setEditMode(v=>!v);cancelForm()}}
              style={{fontSize:10,padding:'2px 8px',borderRadius:5,border:`1px solid ${editMode?C.blue:C.b1}`,background:editMode?'rgba(59,158,237,0.1)':'transparent',color:editMode?C.blue:C.t3,cursor:'pointer',fontFamily:'inherit',fontWeight:600}}>
              {editMode?'Bitti':'Düzenle'}
            </button>
            <button onClick={startAdd}
              style={{fontSize:10,padding:'2px 8px',borderRadius:5,border:`1px solid rgba(16,185,129,0.3)`,background:'rgba(16,185,129,0.07)',color:C.green,cursor:'pointer',fontFamily:'inherit',fontWeight:600}}>
              + Yeni Paket
            </button>
          </>
        )}
        {isFormOpen&&<button onClick={cancelForm} style={{fontSize:10,padding:'2px 8px',borderRadius:5,border:`1px solid ${C.b1}`,background:'transparent',color:C.t3,cursor:'pointer',fontFamily:'inherit'}}>✕ İptal</button>}
      </div>

      {/* Paket butonları */}
      {!isFormOpen&&(
        <div style={{padding:'8px 12px',display:'flex',gap:6,flexWrap:'wrap'}}>
          {presets.map((p,i)=>(
            <div key={i} style={{display:'inline-flex',alignItems:'center',gap:0,borderRadius:7,border:`1px solid ${C.b0}`,overflow:'hidden'}}>
              <button onClick={()=>setST(p.tools)}
                style={{padding:'5px 11px',background:'transparent',border:'none',color:C.t2,cursor:'pointer',fontSize:12,fontFamily:'inherit',fontWeight:600,whiteSpace:'nowrap'}}>
                {p.label}
              </button>
              {editMode&&(
                <>
                  <button onClick={()=>startEdit(i)}
                    style={{padding:'5px 7px',background:'rgba(59,158,237,0.08)',border:'none',borderLeft:`1px solid ${C.b0}`,color:C.blue,cursor:'pointer',fontSize:11,fontFamily:'inherit'}}>✏️</button>
                  <button onClick={()=>deletePreset(i)}
                    style={{padding:'5px 7px',background:'rgba(239,68,68,0.07)',border:'none',borderLeft:`1px solid ${C.b0}`,color:'#f87171',cursor:'pointer',fontSize:11,fontFamily:'inherit'}}>✕</button>
                </>
              )}
            </div>
          ))}
          {presets.length===0&&<span style={{fontSize:11,color:C.t3}}>Henüz paket yok — Yeni Paket ekle</span>}
        </div>
      )}

      {/* Düzenleme / Yeni paket formu */}
      {isFormOpen&&(
        <div style={{padding:'12px 14px'}}>
          <div style={{fontSize:11,fontWeight:700,color:C.t2,marginBottom:10}}>{addMode?'Yeni Paket Oluştur':'Paketi Düzenle'}</div>

          {/* Paket adı */}
          <div style={{marginBottom:10}}>
            <div style={{fontSize:10,color:C.t3,marginBottom:4}}>Paket Adı (emoji + isim)</div>
            <input value={newLabel} onChange={e=>setNewLabel(e.target.value)} placeholder="🚀 Hızlı Başlangıç"
              style={{width:'100%',padding:'7px 10px',borderRadius:6,border:`1px solid ${C.b1}`,background:C.bg2,color:C.t1,fontSize:12,fontFamily:'inherit',outline:'none'}}/>
          </div>

          {/* Araç seçimi */}
          <div style={{fontSize:10,color:C.t3,marginBottom:6}}>Araçları Seç {newTools.length>0&&<span style={{color:C.blue,marginLeft:4}}>{newTools.length} seçili</span>}</div>
          <div style={{maxHeight:200,overflowY:'auto',border:`1px solid ${C.b0}`,borderRadius:7,marginBottom:12}}>
            {allTools.map(t=>{
              const sel = newTools.includes(t.id)
              return(
                <div key={t.id} onClick={()=>toggleTool(t.id)}
                  style={{display:'flex',alignItems:'center',gap:8,padding:'7px 10px',cursor:'pointer',
                    background:sel?'rgba(59,158,237,0.08)':'transparent',
                    borderBottom:`1px solid ${C.b0}`,transition:'background .1s'}}
                  onMouseEnter={e=>!sel&&(e.currentTarget.style.background='rgba(255,255,255,0.02)')}
                  onMouseLeave={e=>!sel&&(e.currentTarget.style.background='transparent')}>
                  <div style={{width:16,height:16,borderRadius:4,border:`1px solid ${sel?C.blue:C.b1}`,background:sel?C.blue:'transparent',flexShrink:0,display:'flex',alignItems:'center',justifyContent:'center'}}>
                    {sel&&<span style={{fontSize:10,color:'#fff',lineHeight:1}}>✓</span>}
                  </div>
                  <span style={{fontSize:13}}>{t.icon}</span>
                  <span style={{fontSize:12,color:sel?C.blue:C.t2,fontWeight:sel?600:400,flex:1}}>{t.label}</span>
                  <span style={{fontSize:10,color:C.t3}}>{t.cat}</span>
                </div>
              )
            })}
          </div>

          <div style={{display:'flex',gap:8}}>
            <BtnPrimary onClick={saveForm} disabled={!newLabel.trim()||!newTools.length}>
              {addMode?'+ Paketi Kaydet':'💾 Güncelle'}
            </BtnPrimary>
            <BtnGhost onClick={cancelForm}>İptal</BtnGhost>
          </div>
        </div>
      )}
    </div>
  )
}

// ─── ACCORDION TOOL LIST ──────────────────────────────────────────────────────
function AccordionToolList({categories, catFilter, filtered, selectedTools, toggle, onDeleteTool, onEditDefault, onMoveUp, onMoveDown}){
  // Tüm araçların kategorilerini al (özel araçlar dahil)
  const allCats = catFilter==='Tümü'
    ? [...new Set([...categories, ...filtered.map(t=>t.cat)])]
    : [catFilter]

  const [open, setOpen] = useState(()=>{
    const init = {}
    allCats.forEach(c=>{ init[c]=false })
    return init
  })
  const toggleCat = cat => setOpen(o=>({...o,[cat]:!o[cat]}))

  return(
    <>
      {allCats.map(cat=>{
        const tools = filtered.filter(t=>t.cat===cat)
        if(!tools.length) return null
        const isOpen = open[cat]
        const selCount = tools.filter(t=>selectedTools.includes(t.id)).length
        return(
          <div key={cat} style={{marginBottom:2}}>
            {/* Kategori başlığı — tıklanabilir */}
            <div onClick={()=>toggleCat(cat)}
              style={{display:'flex',alignItems:'center',gap:6,padding:'7px 8px',borderRadius:7,cursor:'pointer',userSelect:'none',background:isOpen?'rgba(255,255,255,0.03)':'transparent',transition:'background .15s'}}
              onMouseEnter={e=>e.currentTarget.style.background='rgba(255,255,255,0.04)'}
              onMouseLeave={e=>e.currentTarget.style.background=isOpen?'rgba(255,255,255,0.03)':'transparent'}>
              {/* Ok ikonu */}
              <span style={{fontSize:10,color:C.t3,transition:'transform .2s',display:'inline-block',transform:isOpen?'rotate(90deg)':'rotate(0deg)',flexShrink:0}}>▶</span>
              {/* Kategori adı */}
              <span style={{fontSize:11,fontWeight:700,color:isOpen?C.t2:C.t3,textTransform:'uppercase',letterSpacing:.8,flex:1}}>{cat}</span>
              {/* Seçili sayısı */}
              {selCount>0&&<span style={{fontSize:10,fontWeight:700,background:'rgba(59,158,237,0.15)',color:C.blue,padding:'1px 6px',borderRadius:4}}>{selCount}</span>}
              {/* Toplam araç sayısı */}
              <span style={{fontSize:10,color:C.t3}}>{tools.length}</span>
            </div>
            {/* Araçlar — akordiyon */}
            {isOpen&&(
              <div style={{paddingLeft:4,paddingBottom:4}}>
                {tools.map((t,ti)=>(
                  <div key={t.id} style={{display:'flex',alignItems:'center',gap:2}}>
                    <div style={{flex:1}}><ToolChip tool={t} selected={selectedTools.includes(t.id)} onClick={()=>toggle(t.id)} onDelete={onDeleteTool?()=>onDeleteTool(t):null}/></div>
                    <div style={{display:'flex',flexDirection:'column',gap:1,opacity:.4,flexShrink:0}}>
                      <button onClick={()=>onMoveUp&&onMoveUp(t.id)} style={{background:'none',border:'none',color:C.t3,cursor:'pointer',fontSize:9,padding:'1px 3px',lineHeight:1}}>▲</button>
                      <button onClick={()=>onMoveDown&&onMoveDown(t.id)} style={{background:'none',border:'none',color:C.t3,cursor:'pointer',fontSize:9,padding:'1px 3px',lineHeight:1}}>▼</button>
                    </div>
                    {onEditDefault&&<button onClick={()=>onEditDefault(t)} style={{fontSize:10,padding:'2px 6px',borderRadius:4,border:`1px solid ${C.b0}`,background:'transparent',color:C.t3,cursor:'pointer',fontFamily:'inherit',flexShrink:0}}>✏️</button>}
                  </div>
                ))}
              </div>
            )}
          </div>
        )
      })}
    </>
  )
}


// ─── ARAÇ TİPLERİ ─────────────────────────────────────────────────────────────
const TOOL_TYPES = {
  // ── Analiz & Chat ──────────────────────────────────
  chat:          { label:'Chat / Analiz',         icon:'💬', group:'Analiz',      desc:'Sohbet, analiz, senaryo üretimi',           color:'#3b9eed', outputLabel:'Metin çıktısı' },
  // ── Web Otomasyon ──────────────────────────────────
  playwright:    { label:'Playwright (TS)',        icon:'🎭', group:'Web',         desc:'JS/TS Playwright E2E kodu',                 color:'#10b981', outputLabel:'.spec.ts', ext:'spec.ts' },
  selenium_py:   { label:'Selenium Python',        icon:'🐍', group:'Web',         desc:'Python Selenium web otomasyon kodu',        color:'#facc15', outputLabel:'.py',       ext:'py' },
  selenium_java: { label:'Selenium Java',          icon:'☕', group:'Web',         desc:'Java Selenium test kodu',                   color:'#f97316', outputLabel:'.java',     ext:'java' },
  cypress:       { label:'Cypress',                icon:'🌲', group:'Web',         desc:'Cypress E2E test kodu',                     color:'#06b6d4', outputLabel:'.cy.js',    ext:'cy.js' },
  puppeteer:     { label:'Puppeteer',              icon:'🤖', group:'Web',         desc:'Node.js Puppeteer otomasyon kodu',          color:'#84cc16', outputLabel:'.js',       ext:'js' },
  // ── Mobil ──────────────────────────────────────────
  appium:        { label:'Appium (Mobil)',          icon:'📱', group:'Mobil',       desc:'iOS/Android Appium test kodu',              color:'#8b5cf6', outputLabel:'.py',       ext:'py' },
  espresso:      { label:'Espresso (Android)',      icon:'🤖', group:'Mobil',       desc:'Android native Espresso test kodu',         color:'#34d399', outputLabel:'.kt',       ext:'kt' },
  xcuitest:      { label:'XCUITest (iOS)',          icon:'🍎', group:'Mobil',       desc:'iOS native XCUITest kodu',                  color:'#94a3b8', outputLabel:'.swift',    ext:'swift' },
  // ── API & Servis ───────────────────────────────────
  postman:       { label:'Postman Collection',      icon:'📮', group:'API',         desc:'Postman JSON koleksiyonu',                  color:'#f97316', outputLabel:'.json',     ext:'json' },
  rest_assured:  { label:'REST Assured (Java)',     icon:'🔌', group:'API',         desc:'Java REST Assured API test kodu',           color:'#6366f1', outputLabel:'.java',     ext:'java' },
  pytest_api:    { label:'pytest (Python)',         icon:'🧪', group:'API',         desc:'Python pytest API test kodu',               color:'#f43f5e', outputLabel:'.py',       ext:'py' },
  graphql:       { label:'GraphQL Test',            icon:'🔮', group:'API',         desc:'GraphQL sorgu/mutasyon testleri',           color:'#ec4899', outputLabel:'.js',       ext:'js' },
  // ── Performans ─────────────────────────────────────
  jmeter:        { label:'JMeter',                  icon:'⚡', group:'Performans',  desc:'JMeter .jmx yük testi planı',               color:'#f59e0b', outputLabel:'.js',       ext:'js' },
  k6:            { label:'k6 (Performans)',          icon:'📈', group:'Performans',  desc:'k6 JavaScript yük testi scripti',           color:'#84cc16', outputLabel:'.js',       ext:'js' },
  gatling:       { label:'Gatling (Scala)',          icon:'🏹', group:'Performans',  desc:'Gatling Scala performans testi',            color:'#06b6d4', outputLabel:'.scala',    ext:'scala' },
  // ── Güvenlik ───────────────────────────────────────
  owasp_zap:     { label:'OWASP ZAP Script',        icon:'🛡', group:'Güvenlik',    desc:'OWASP ZAP güvenlik test scripti',           color:'#ef4444', outputLabel:'.py',       ext:'py' },
  burp_suite:    { label:'Burp Suite',              icon:'🔐', group:'Güvenlik',    desc:'Burp Suite test senaryosu',                 color:'#dc2626', outputLabel:'.txt',      ext:'txt' },
  // ── CI/CD & DevOps ─────────────────────────────────
  github_actions:{ label:'GitHub Actions',          icon:'🔄', group:'DevOps',      desc:'GitHub Actions CI pipeline YAML',           color:'#3b9eed', outputLabel:'.yml',      ext:'yml' },
  jenkins:       { label:'Jenkins Pipeline',        icon:'🏗',  group:'DevOps',      desc:'Jenkinsfile pipeline scripti',              color:'#94a3b8', outputLabel:'.groovy',   ext:'groovy' },
  docker:        { label:'Docker / K8s',            icon:'🐳', group:'DevOps',      desc:'Docker compose veya K8s test config',       color:'#38bdf8', outputLabel:'.yml',      ext:'yml' },
  // ── Veritabanı ─────────────────────────────────────
  sql_test:      { label:'SQL Test Sorguları',      icon:'🗄', group:'Veritabanı',  desc:'Veritabanı doğrulama SQL sorguları',        color:'#6366f1', outputLabel:'.sql',      ext:'sql' },
  // ── Masaüstü ───────────────────────────────────────
  winappdriver:  { label:'WinAppDriver',            icon:'🖥️', group:'Masaüstü',   desc:'Windows uygulama otomasyon kodu',           color:'#3b82f6', outputLabel:'.cs',       ext:'cs' },
  pywinauto:     { label:'pywinauto (Python)',      icon:'🪟', group:'Masaüstü',   desc:'Python pywinauto masaüstü testi',           color:'#facc15', outputLabel:'.py',       ext:'py' },
  // ── Raporlama ──────────────────────────────────────
  html_report:   { label:'HTML Test Raporu',        icon:'📊', group:'Rapor',       desc:'Görsel HTML test raporu üret',              color:'#ec4899', outputLabel:'.html',     ext:'html' },
  allure:        { label:'Allure Report',           icon:'🌸', group:'Rapor',       desc:'Allure test rapor konfigürasyonu',          color:'#f472b6', outputLabel:'.json',     ext:'json' },
  // ── Diğer ──────────────────────────────────────────
  gherkin:       { label:'BDD / Gherkin',           icon:'🥒', group:'Doküman',     desc:'Cucumber/Behave .feature dosyası',          color:'#34d399', outputLabel:'.feature',  ext:'feature' },
  batch:         { label:'Batch / Shell',           icon:'💻', group:'Script',      desc:'Windows .bat veya Linux .sh scripti',       color:'#94a3b8', outputLabel:'.bat',      ext:'bat' },
  powershell:    { label:'PowerShell',              icon:'💙', group:'Script',      desc:'PowerShell test/otomasyon scripti',         color:'#3b82f6', outputLabel:'.ps1',      ext:'ps1' },
}

// Kod/dosya çıktısı veren tipler
const CODE_TYPES = Object.keys(TOOL_TYPES).filter(k=>k!=='chat')

// Çıktıyı dosya olarak indir
function downloadFile(content, filename, mime='text/plain') {
  const blob = new Blob([content], {type: mime+';charset=utf-8'})
  const a = Object.assign(document.createElement('a'), {
    href: URL.createObjectURL(blob),
    download: filename
  })
  a.click()
  URL.revokeObjectURL(a.href)
}

// Kod bloğunu çıkart
function extractCode(text) {
  const match = text.match(/```[\w]*\n?([\s\S]*?)```/)
  return match ? match[1].trim() : text.trim()
}

// ─── CUSTOM TOOL ROW ──────────────────────────────────────────────────────────
function CustomToolRow({tool, isEditing, onEdit, onDelete}){
  const [h,sH] = useState(false)
  const cl = PALETTE[tool.color]||PALETTE.blue
  const tt = TOOL_TYPES[tool.toolType]
  return(
    <div onMouseEnter={()=>sH(true)} onMouseLeave={()=>sH(false)}
      style={{position:'relative',marginBottom:3}}>
      <div style={{display:'flex',alignItems:'center',gap:9,padding:'8px 10px',borderRadius:8,
        border:`1px solid ${isEditing?C.blue:h?C.b1:C.b0}`,
        background:isEditing?'rgba(59,158,237,0.08)':h?'rgba(255,255,255,0.02)':'transparent',
        transition:'all .15s'}}>
        <span style={{fontSize:16,lineHeight:1,opacity:.8,flexShrink:0}}>{tool.icon}</span>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:12,fontWeight:600,color:isEditing?C.blue:C.t1,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{tool.label}</div>
          <div style={{display:'flex',alignItems:'center',gap:5,marginTop:2}}>
            {tt&&<span style={{fontSize:10,padding:'1px 6px',borderRadius:4,background:`${tt.color}18`,color:tt.color,fontWeight:600,border:`1px solid ${tt.color}30`}}>{tt.icon} {tt.label}</span>}
            <span style={{fontSize:10,color:C.t3}}>{tool.desc}</span>
          </div>
        </div>
        {/* Hover aksiyonlar */}
        {(h||isEditing)&&(
          <div style={{display:'flex',gap:4,flexShrink:0}}>
            <button onClick={onEdit} style={{fontSize:11,padding:'3px 9px',borderRadius:5,
              border:`1px solid ${isEditing?C.blue:'rgba(59,158,237,0.3)'}`,
              background:isEditing?'rgba(59,158,237,0.15)':'rgba(59,158,237,0.08)',
              color:C.blue,cursor:'pointer',fontFamily:'inherit',fontWeight:600}}>
              {isEditing?'Düzenleniyor':'✏️ Düzenle'}
            </button>
            <button onClick={onDelete} style={{fontSize:11,padding:'3px 7px',borderRadius:5,
              border:'1px solid rgba(239,68,68,0.2)',background:'rgba(239,68,68,0.07)',
              color:'#f87171',cursor:'pointer',fontFamily:'inherit'}}>✕</button>
          </div>
        )}
      </div>
    </div>
  )
}

// ─── TOOL CHIP ────────────────────────────────────────────────────────────────

// ─── CONFIRM DIALOG ───────────────────────────────────────────────────────────
function ConfirmDialog({message, onConfirm, onCancel}){
  return(
    <div style={{position:'fixed',inset:0,zIndex:9999,display:'flex',alignItems:'center',justifyContent:'center',background:'rgba(0,0,0,0.6)',backdropFilter:'blur(4px)'}}>
      <div style={{background:C.bg2,border:`1px solid ${C.b1}`,borderRadius:12,padding:'24px 28px',maxWidth:360,width:'90%',boxShadow:'0 20px 60px rgba(0,0,0,0.5)'}}>
        <div style={{fontSize:20,marginBottom:12,textAlign:'center'}}>🗑️</div>
        <div style={{fontSize:13,color:C.t1,textAlign:'center',lineHeight:1.6,marginBottom:20}}>{message}</div>
        <div style={{display:'flex',gap:8,justifyContent:'center'}}>
          <BtnDanger onClick={onConfirm} icon="✕">Kaldır</BtnDanger>
          <BtnGhost onClick={onCancel}>İptal</BtnGhost>
        </div>
      </div>
    </div>
  )
}

function ToolChip({tool, selected, onClick, onDelete}){
  const cl = PALETTE[tool.color]||PALETTE.blue
  const [h,sH]=useState(false)
  return(
    <div onMouseEnter={()=>sH(true)} onMouseLeave={()=>sH(false)}
      style={{display:'flex',alignItems:'center',gap:9,padding:'8px 10px',borderRadius:8,border:`1px solid ${selected?cl.border:h?C.b1:C.b0}`,background:selected?cl.sel:h?'rgba(255,255,255,0.02)':'transparent',transition:'all .15s',cursor:'pointer',marginBottom:2}}>
      <span style={{flex:1,display:'flex',alignItems:'center',gap:9,minWidth:0}} onClick={onClick}>
        <span style={{fontSize:16,lineHeight:1,opacity:selected?1:.65,flexShrink:0}}>{tool.icon}</span>
        <div style={{minWidth:0}}>
          <div style={{fontSize:12,fontWeight:600,color:selected?cl.text:C.t1,lineHeight:1.3,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{tool.label}</div>
          <div style={{fontSize:11,color:C.t3,lineHeight:1.3,marginTop:1,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{tool.desc}</div>
        </div>
      </span>
      {selected&&<span style={{fontSize:10,color:cl.text,flexShrink:0,background:`${cl.text}20`,padding:'2px 6px',borderRadius:4}}>✓</span>}
      {onDelete&&h&&<button onClick={e=>{e.stopPropagation();onDelete()}} style={{background:'rgba(239,68,68,0.08)',border:'1px solid rgba(239,68,68,0.2)',borderRadius:4,color:'#f87171',cursor:'pointer',fontSize:11,padding:'1px 5px',flexShrink:0,lineHeight:1.4,fontFamily:'inherit'}}
        onMouseEnter={e=>{e.currentTarget.style.background='rgba(239,68,68,0.18)'}}
        onMouseLeave={e=>{e.currentTarget.style.background='rgba(239,68,68,0.08)'}}>✕</button>}
    </div>
  )
}

// ─── RESULT CARD ──────────────────────────────────────────────────────────────
const EXPORT_FORMATS = [
  { value:'txt',  label:'📄 Text (.txt)' },
  { value:'html', label:'🌐 HTML (.html)' },
  { value:'md',   label:'📝 Markdown (.md)' },
  { value:'csv',  label:'📊 CSV (.csv)' },
  { value:'json', label:'🔧 JSON (.json)' },
]

function exportContent(content, fmt, label) {
  const base = (label||'sonuc').replace(/\s+/g,'_').toLowerCase()
  if(fmt==='txt') { downloadFile(content, base+'.txt', 'text/plain'); return }
  if(fmt==='md')  { downloadFile(content, base+'.md',  'text/markdown'); return }
  if(fmt==='csv') {
    const rows = content.split('\n').map(l=>'"'+l.replace(/"/g,'""')+'"').join('\n')
    downloadFile(rows, base+'.csv', 'text/csv'); return
  }
  if(fmt==='json') {
    const obj = JSON.stringify({tool:label, content, date:new Date().toISOString()}, null, 2)
    downloadFile(obj, base+'.json', 'application/json'); return
  }
  if(fmt==='html') { downloadBeniOku({label,toolType:null}, content); return }
}

function ResultCard({tool, content, messages, onClose, onMessagesUpdate}){
  const cl = PALETTE[tool.color]||PALETTE.blue
  const [copied,sC]=useState(false)
  const [open,sO]=useState(true)
  const [chatOpen,setChatOpen]=useState(false)
  const [chatInput,setChatInput]=useState('')
  const [chatLoading,setChatLoading]=useState(false)
  const [exportFmt,setExportFmt]=useState('txt')
  const [msgs,setMsgs]=useState(messages||[{role:'assistant',content}])
  const chatBottom=useRef()
  const tt = tool.toolType ? TOOL_TYPES[tool.toolType] : null
  const isCode = tool.toolType && CODE_TYPES.includes(tool.toolType)

  useEffect(()=>{ chatBottom.current?.scrollIntoView({behavior:'smooth'}) },[msgs,chatLoading])

  const downloadCode=()=>{
    const code = extractCode(content)
    const ext = tt?.ext||'txt'
    const fname = (tool.label||'output').replace(/\s+/g,'_').toLowerCase()+'.'+ext
    const mime = ext==='html'?'text/html':ext==='py'?'text/x-python':'text/plain'
    downloadFile(code, fname, mime)
  }

  const sendChat=async()=>{
    if(!chatInput.trim()||chatLoading) return
    const userMsg = {role:'user', content:chatInput}
    const newMsgs = [...msgs, userMsg]
    setMsgs(newMsgs)
    setChatInput('')
    setChatLoading(true)
    try{
      // Araç bağlamını sistem mesajı olarak ekle
      const systemCtx = {role:'user', content:`Sen "${tool.label}" aracının sonucunu derinleştiriyorsun. Orijinal sonuç:\n\n${content}\n\n---\nBu bağlamda devam et:`}
      const apiMsgs = [systemCtx, ...newMsgs.slice(-6)]
      const reply = await callAI(apiMsgs)
      const updated = [...newMsgs, {role:'assistant', content:reply}]
      setMsgs(updated)
      onMessagesUpdate&&onMessagesUpdate(updated)
    }catch(e){ setMsgs(m=>[...m,{role:'assistant',content:'Hata: '+e.message}]) }
    finally{ setChatLoading(false) }
  }

  return(
    <div style={{border:`1px solid ${cl.border}`,borderRadius:10,overflow:'hidden',marginBottom:10}}>
      {/* Header */}
      <div onClick={()=>sO(v=>!v)} style={{display:'flex',alignItems:'center',gap:8,padding:'10px 14px',background:cl.bg,cursor:'pointer',userSelect:'none',flexWrap:'wrap'}}>
        <span style={{fontSize:15}}>{tool.icon}</span>
        <span style={{fontWeight:700,color:cl.text,fontSize:12,flex:1,minWidth:80}}>{tool.label}</span>
        {tt&&<span style={{fontSize:10,padding:'2px 7px',borderRadius:4,background:`${tt.color}22`,color:tt.color,fontWeight:600,border:`1px solid ${tt.color}40`,flexShrink:0}}>{tt.icon} {tt.label}</span>}
        {/* Format seçici + indir */}
        <div style={{display:'flex',alignItems:'center',gap:4}} onClick={e=>e.stopPropagation()}>
          <select value={exportFmt} onChange={e=>setExportFmt(e.target.value)}
            style={{fontSize:11,padding:'3px 6px',borderRadius:5,border:`1px solid ${C.b1}`,background:C.bg3,color:C.t2,fontFamily:'inherit',cursor:'pointer'}}>
            {EXPORT_FORMATS.map(f=><option key={f.value} value={f.value}>{f.label}</option>)}
          </select>
          <BtnSuccess size="sm" icon="📥" onClick={()=>exportContent(content,exportFmt,tool.label)}>İndir</BtnSuccess>
        </div>
        {isCode&&<BtnSuccess size="sm" onClick={e=>{e.stopPropagation();downloadCode()}} icon="💾">{tt?.ext?.toUpperCase()}</BtnSuccess>}
        <BtnGhost size="sm" icon="📄" onClick={e=>{e.stopPropagation();downloadBeniOku(tool,content)}}>Beni Oku</BtnGhost>
        <BtnGhost size="sm" onClick={e=>{e.stopPropagation();navigator.clipboard.writeText(content);sC(true);setTimeout(()=>sC(false),2000)}} color={copied?C.green:cl.text}>{copied?'✓':'Kopyala'}</BtnGhost>
        <button onClick={e=>{e.stopPropagation();onClose()}} style={{background:'none',border:'none',color:C.t3,cursor:'pointer',fontSize:14,padding:'0 2px'}} onMouseEnter={e=>e.currentTarget.style.color=C.t1} onMouseLeave={e=>e.currentTarget.style.color=C.t3}>✕</button>
        <span style={{color:C.t3,fontSize:10}}>{open?'▲':'▼'}</span>
      </div>

      {/* İlk sonuç */}
      {open&&(
        <div style={{padding:'14px 16px',background:'rgba(0,0,0,0.18)'}}>
          <pre style={{whiteSpace:'pre-wrap',wordBreak:'break-word',fontSize:13,color:C.t1,lineHeight:1.8,fontFamily:"'Segoe UI',system-ui,sans-serif",margin:0}}>{content}</pre>
        </div>
      )}

      {/* Devam sohbeti */}
      {open&&(
        <div style={{borderTop:`1px solid ${C.b0}`}}>
          {/* Sohbet toggle */}
          <div onClick={()=>setChatOpen(v=>!v)}
            style={{display:'flex',alignItems:'center',gap:6,padding:'7px 14px',cursor:'pointer',background:'rgba(255,255,255,0.01)',userSelect:'none'}}
            onMouseEnter={e=>e.currentTarget.style.background='rgba(255,255,255,0.03)'}
            onMouseLeave={e=>e.currentTarget.style.background='rgba(255,255,255,0.01)'}>
            <span style={{fontSize:12}}>💬</span>
            <span style={{fontSize:11,color:C.t3,flex:1}}>
              {msgs.length>1 ? `Devam sohbeti (${msgs.length-1} mesaj)` : 'Bu sonuç hakkında soru sor veya derinleştir'}
            </span>
            <span style={{fontSize:10,color:C.t3}}>{chatOpen?'▲':'▼'}</span>
          </div>

          {/* Sohbet alanı */}
          {chatOpen&&(
            <div style={{padding:'10px 14px',background:'rgba(0,0,0,0.12)'}}>
              {/* Geçmiş mesajlar (ilk asistan mesajı hariç) */}
              {msgs.slice(1).map((m,i)=>(
                <div key={i} style={{marginBottom:10,display:'flex',gap:8,flexDirection:m.role==='user'?'row-reverse':'row'}}>
                  <div style={{width:24,height:24,borderRadius:'50%',flexShrink:0,display:'flex',alignItems:'center',justifyContent:'center',fontSize:12,
                    background:m.role==='user'?'rgba(59,158,237,0.2)':'rgba(139,92,246,0.2)'}}>
                    {m.role==='user'?'👤':'🤖'}
                  </div>
                  <div style={{maxWidth:'85%',padding:'8px 12px',borderRadius:8,fontSize:12,lineHeight:1.7,
                    background:m.role==='user'?'rgba(59,158,237,0.1)':'rgba(255,255,255,0.04)',
                    border:`1px solid ${m.role==='user'?'rgba(59,158,237,0.2)':C.b0}`,
                    color:C.t1,whiteSpace:'pre-wrap',wordBreak:'break-word'}}>
                    {m.content}
                  </div>
                </div>
              ))}
              {chatLoading&&(
                <div style={{display:'flex',gap:8,marginBottom:10}}>
                  <div style={{width:24,height:24,borderRadius:'50%',background:'rgba(139,92,246,0.2)',display:'flex',alignItems:'center',justifyContent:'center',fontSize:12}}>🤖</div>
                  <div style={{padding:'8px 12px',borderRadius:8,background:'rgba(255,255,255,0.04)',border:`1px solid ${C.b0}`}}>
                    <Spin c={C.purple} s={12}/>
                  </div>
                </div>
              )}
              <div ref={chatBottom}/>
              {/* Input */}
              <div style={{display:'flex',gap:6,marginTop:8}}>
                <input value={chatInput} onChange={e=>setChatInput(e.target.value)}
                  onKeyDown={e=>e.key==='Enter'&&!e.shiftKey&&sendChat()}
                  placeholder="Devam et, soru sor veya değiştir..."
                  style={{flex:1,padding:'8px 12px',borderRadius:7,border:`1px solid ${C.b1}`,background:C.bg2,color:C.t1,fontSize:12,fontFamily:'inherit',outline:'none'}}/>
                <BtnAccent size="sm" onClick={sendChat} loading={chatLoading} disabled={!chatInput.trim()}>Gönder</BtnAccent>
              </div>
              <div style={{fontSize:10,color:C.t3,marginTop:5}}>Enter ile gönder · Shift+Enter yeni satır</div>
            </div>
          )}
        </div>
      )}
    </div>
  )
}

// ─── ICD PANEL ────────────────────────────────────────────────────────────────
function ICDPanel({icdContent, setIcdContent}){
  const [loading,sL]=useState(false)
  const ref=useRef()
  const handleFile=async(e)=>{
    const f=e.target.files[0]; if(!f) return
    sL(true)
    try{
      if(f.name.endsWith('.docx')){
        const m=await import('mammoth').catch(()=>null)
        if(m){ const r=await m.extractRawText({arrayBuffer:await f.arrayBuffer()}); setIcdContent(r.value); LS.set('qa_icd_content',r.value) }
      } else { const t=await f.text(); setIcdContent(t); LS.set('qa_icd_content',t) }
    }catch(e){alert('Dosya okunamadı: '+e.message)}
    finally{sL(false)}
    e.target.value=''
  }
  if(!icdContent) return(
    <div style={{display:'flex',alignItems:'center',gap:8,padding:'8px 12px',borderRadius:8,border:`1px dashed ${C.b1}`,background:'rgba(255,255,255,0.01)',marginBottom:12,cursor:'pointer'}} onClick={()=>ref.current.click()}>
      <span style={{fontSize:13}}>📎</span>
      <span style={{fontSize:12,color:C.t3,flex:1}}>ICD/Doküman yükle (TXT, MD, DOCX) — test üretiminde baz alınır</span>
      <BtnGhost size="sm">{loading?<Spin s={12}/>:'Dosya Seç'}</BtnGhost>
      <input ref={ref} type="file" accept=".txt,.md,.docx" style={{display:'none'}} onChange={handleFile}/>
    </div>
  )
  return(
    <div style={{display:'flex',alignItems:'center',gap:8,padding:'8px 12px',borderRadius:8,border:`1px solid ${C.b2}`,background:'rgba(59,158,237,0.04)',marginBottom:12}}>
      <span style={{fontSize:13}}>📎</span>
      <Tag color={C.green}>ICD Yüklendi</Tag>
      <span style={{fontSize:11,color:C.t3,flex:1}}>{icdContent.slice(0,60)}...</span>
      <BtnGhost size="sm" onClick={()=>ref.current.click()}>{loading?<Spin s={12}/>:'Değiştir'}</BtnGhost>
      <BtnDanger size="sm" onClick={()=>{setIcdContent('');LS.set('qa_icd_content','')}}>✕</BtnDanger>
      <input ref={ref} type="file" accept=".txt,.md,.docx" style={{display:'none'}} onChange={handleFile}/>
    </div>
  )
}

// ─── DASHBOARD ────────────────────────────────────────────────────────────────
function DashboardTab({usageStats, allTools, bugs, setBugs}){
  const [showForm,sF]=useState(false)
  const [form,setForm]=useState({title:'',version:'',severity:'High',status:'Open',desc:''})
  const [statusFilter,sSF]=useState('all')
  const [importMsg,sIM]=useState(null)
  const csvRef=useRef()

  const importCSV=async(e)=>{
    const file=e.target.files[0]; if(!file) return
    try{
      const text=await file.text()
      const lines=text.split('\n').filter(l=>l.trim())
      const heads=lines[0].split(',').map(h=>h.replace(/"/g,'').trim().toLowerCase())
      const col=(row,...ns)=>{ const cs=row.match(/(".*?"|[^,]+)/g)||[]; for(const n of ns){const i=heads.findIndex(h=>h.includes(n));if(i>=0)return(cs[i]||'').replace(/"/g,'').trim()}; return'' }
      const sm={highest:'Highest',critical:'Highest',high:'High',medium:'Medium',low:'Low',lowest:'Lowest',blocker:'Highest',major:'High',minor:'Low',trivial:'Lowest'}
      const stm={open:'Open','to do':'Open','in progress':'In Progress','in review':'In Progress',done:'Closed',resolved:'Resolved',closed:'Closed'}
      const imp=lines.slice(1).filter(l=>l.trim()).map(r=>({id:Date.now()+Math.random(),title:col(r,'summary','title','issue')||'İsimsiz',severity:sm[col(r,'priority','severity').toLowerCase()]||'Medium',status:stm[col(r,'status').toLowerCase()]||'Open',version:(col(r,'fix version','version','affects')||'?').split(',')[0].trim()||'?',date:col(r,'created','date')?.slice(0,10)||new Date().toISOString().slice(0,10),desc:col(r,'description','desc')||''}))
      const upd=[...bugs,...imp]; setBugs(upd); LS.set('qa_bugs',upd)
      sIM(`✓ ${imp.length} bug aktarıldı`); setTimeout(()=>sIM(null),3000)
    }catch(er){sIM('Hata: '+er.message);setTimeout(()=>sIM(null),4000)}
    e.target.value=''
  }

  const addBug=()=>{
    if(!form.title.trim()||!form.version.trim()) return
    const upd=[...bugs,{...form,id:Date.now(),date:new Date().toISOString().slice(0,10)}]
    setBugs(upd); LS.set('qa_bugs',upd)
    setForm({title:'',version:'',severity:'High',status:'Open',desc:''}); sF(false)
  }
  const delBug=id=>{ const u=bugs.filter(b=>b.id!==id); setBugs(u); LS.set('qa_bugs',u) }
  const updBug=(id,k,v)=>{ const u=bugs.map(b=>b.id===id?{...b,[k]:v}:b); setBugs(u); LS.set('qa_bugs',u) }

  const total=bugs.length
  const openCount=bugs.filter(b=>b.status==='Open').length
  const highCount=bugs.filter(b=>b.status!=='Closed'&&b.status!=='Resolved'&&(b.severity==='Highest'||b.severity==='High')).length
  const closedCount=bugs.filter(b=>b.status==='Closed'||b.status==='Resolved').length
  const toolTotal=Object.values(usageStats).reduce((a,b)=>a+b,0)

  const byVersion={}; bugs.forEach(b=>{byVersion[b.version]=(byVersion[b.version]||0)+1})
  const versions=Object.entries(byVersion).sort((a,b)=>a[0].localeCompare(b[0]))
  const maxVer=Math.max(1,...versions.map(v=>v[1]))
  const topTools=Object.entries(usageStats).sort((a,b)=>b[1]-a[1]).slice(0,5)
  const maxUse=topTools[0]?.[1]||1

  const filtered=statusFilter==='all'?bugs:bugs.filter(b=>b.status===statusFilter||b.severity===statusFilter)

  return(
    <div style={{height:'100%',overflowY:'auto',background:C.bg0}}>
      <div style={{padding:'20px 24px',maxWidth:1200,margin:'0 auto'}}>

        {/* Stat row */}
        <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(180px,1fr))',gap:12,marginBottom:20}}>
          {[
            {icon:'🐛',label:'Toplam Bug',val:total,color:C.red,sub:`${openCount} açık`},
            {icon:'🔴',label:'Aktif High+',val:highCount,color:'#ff4040',sub:'çözülmemiş'},
            {icon:'✅',label:'Çözülen',val:closedCount,color:C.green,sub:`${total?Math.round(closedCount/total*100):0}% tamamlandı`},
            {icon:'⚡',label:'Araç Kullanımı',val:toolTotal,color:C.purple,sub:'toplam çalıştırma'},
          ].map((s,i)=>(
            <div key={i} style={{background:C.bg2,border:`1px solid ${C.b0}`,borderRadius:12,padding:'16px 18px',position:'relative',overflow:'hidden'}}>
              <div style={{position:'absolute',top:-10,right:-8,fontSize:44,opacity:.06}}>{s.icon}</div>
              <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:10}}>
                <div style={{width:8,height:8,borderRadius:'50%',background:s.color,boxShadow:`0 0 6px ${s.color}`}}/>
                <span style={{fontSize:11,fontWeight:600,color:C.t3,textTransform:'uppercase',letterSpacing:.6}}>{s.label}</span>
              </div>
              <div style={{fontSize:28,fontWeight:900,color:s.color,lineHeight:1,letterSpacing:-1}}>{s.val}</div>
              <div style={{fontSize:11,color:C.t3,marginTop:5}}>{s.sub}</div>
            </div>
          ))}
        </div>

        {/* Bug tracker */}
        <div style={{background:C.bg2,border:`1px solid ${C.b0}`,borderRadius:12,marginBottom:20,overflow:'hidden'}}>
          {/* Header */}
          <div style={{display:'flex',alignItems:'center',gap:10,padding:'14px 18px',borderBottom:`1px solid ${C.b0}`,flexWrap:'wrap',gap:8}}>
            <span style={{fontSize:13,fontWeight:700,color:C.t0,flex:1}}>🐛 Bug Takibi</span>
            {bugs.length>0&&<DashExport bugs={bugs}/>}
            {importMsg&&<Tag color={importMsg.startsWith('✓')?C.green:C.red}>{importMsg}</Tag>}
            {/* Status filtreler */}
            <div style={{display:'flex',gap:4,flexWrap:'wrap'}}>
              {['all','Open','In Progress','Resolved','Closed'].map(f=>(
                <button key={f} onClick={()=>sSF(f)} style={{padding:'4px 10px',borderRadius:6,border:`1px solid ${statusFilter===f?C.b2:C.b0}`,background:statusFilter===f?'rgba(59,158,237,0.1)':'transparent',color:statusFilter===f?C.blue:C.t3,cursor:'pointer',fontSize:11,fontWeight:600,fontFamily:'inherit',transition:'all .15s'}}>
                  {f==='all'?'Tümü':f}
                </button>
              ))}
            </div>
            <Divider sx={{width:1,height:18,background:C.b1}}/>
            <BtnSuccess size="sm" icon="📥" onClick={()=>csvRef.current.click()}>JIRA CSV</BtnSuccess>
            <input ref={csvRef} type="file" accept=".csv" style={{display:'none'}} onChange={importCSV}/>
            <BtnPrimary size="sm" icon={showForm?'':'+'} onClick={()=>sF(v=>!v)}>{showForm?'İptal':'Bug Ekle'}</BtnPrimary>
          </div>

          {/* Add form */}
          {showForm&&(
            <div style={{padding:'16px 18px',borderBottom:`1px solid ${C.b0}`,background:'rgba(255,255,255,0.015)'}}>
              <div style={{display:'grid',gridTemplateColumns:'2fr 1fr',gap:10,marginBottom:10}}>
                <TextInput value={form.title} onChange={e=>setForm(f=>({...f,title:e.target.value}))} placeholder="Bug başlığı *"/>
                <TextInput value={form.version} onChange={e=>setForm(f=>({...f,version:e.target.value}))} placeholder="Versiyon (örn: 2.1.0) *"/>
              </div>
              <div style={{display:'grid',gridTemplateColumns:'1fr 1fr 2fr',gap:10,marginBottom:12}}>
                <Sel value={form.severity} onChange={e=>setForm(f=>({...f,severity:e.target.value}))}>
                  {Object.keys(SEV).map(s=><option key={s}>{s}</option>)}
                </Sel>
                <Sel value={form.status} onChange={e=>setForm(f=>({...f,status:e.target.value}))}>
                  {Object.keys(STA).map(s=><option key={s}>{s}</option>)}
                </Sel>
                <TextInput value={form.desc} onChange={e=>setForm(f=>({...f,desc:e.target.value}))} placeholder="Kısa açıklama (opsiyonel)"/>
              </div>
              <div style={{display:'flex',gap:8}}>
                <BtnPrimary onClick={addBug} disabled={!form.title.trim()||!form.version.trim()}>✓ Kaydet</BtnPrimary>
                <BtnGhost onClick={()=>sF(false)}>İptal</BtnGhost>
              </div>
            </div>
          )}

          {/* Bug table */}
          {filtered.length===0?(
            <div style={{padding:'36px 0',textAlign:'center',color:C.t3}}>
              <div style={{fontSize:32,marginBottom:8,opacity:.4}}>🎉</div>
              <div style={{fontSize:13}}>{bugs.length===0?'Henüz bug eklenmedi':'Bu filtreye uygun bug yok'}</div>
            </div>
          ):(
            <>
              <div style={{display:'grid',gridTemplateColumns:'1fr 100px 130px 140px 90px 32px',gap:8,padding:'7px 18px',borderBottom:`1px solid ${C.b0}`,fontSize:10,fontWeight:700,color:C.t3,textTransform:'uppercase',letterSpacing:.7}}>
                <span>Başlık / Açıklama</span><span>Versiyon</span><span>Severity</span><span>Durum</span><span>Tarih</span><span/>
              </div>
              {filtered.map((b,i)=>{
                const sc=SEV[b.severity]||SEV.Medium
                const stc=STA[b.status]||STA.Open
                return(
                  <div key={b.id} style={{display:'grid',gridTemplateColumns:'1fr 100px 130px 140px 90px 32px',gap:8,padding:'10px 18px',borderBottom:i<filtered.length-1?`1px solid ${C.b0}`:'none',alignItems:'center',transition:'background .1s'}}
                    onMouseEnter={e=>e.currentTarget.style.background='rgba(255,255,255,0.01)'}
                    onMouseLeave={e=>e.currentTarget.style.background='transparent'}>
                    <div>
                      <div style={{fontSize:13,color:C.t1,fontWeight:500}}>{b.title}</div>
                      {b.desc&&<div style={{fontSize:11,color:C.t3,marginTop:2}}>{b.desc.slice(0,55)}{b.desc.length>55?'…':''}</div>}
                    </div>
                    <span style={{fontSize:12,color:C.t3,fontFamily:'monospace'}}>v{b.version}</span>
                    <select value={b.severity} onChange={e=>updBug(b.id,'severity',e.target.value)}
                      style={{padding:'4px 7px',fontSize:11,fontWeight:600,background:sc.bg,color:sc.fg,border:`1px solid ${sc.ring}`,borderRadius:6,fontFamily:'inherit',outline:'none',cursor:'pointer'}}>
                      {Object.keys(SEV).map(s=><option key={s}>{s}</option>)}
                    </select>
                    <select value={b.status} onChange={e=>updBug(b.id,'status',e.target.value)}
                      style={{padding:'4px 7px',fontSize:11,fontWeight:600,background:stc.bg,color:stc.fg,border:`1px solid ${stc.fg}44`,borderRadius:6,fontFamily:'inherit',outline:'none',cursor:'pointer'}}>
                      {Object.keys(STA).map(s=><option key={s}>{s}</option>)}
                    </select>
                    <span style={{fontSize:11,color:C.t3}}>{b.date}</span>
                    <button onClick={()=>delBug(b.id)} style={{background:'none',border:'none',color:C.t3,cursor:'pointer',fontSize:13,padding:0,lineHeight:1}}
                      onMouseEnter={e=>e.currentTarget.style.color='#f87171'}
                      onMouseLeave={e=>e.currentTarget.style.color=C.t3}>✕</button>
                  </div>
                )
              })}
            </>
          )}
        </div>

        {/* Charts row */}
        <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:16,marginBottom:16}}>
          {/* Severity */}
          <div style={{background:C.bg2,border:`1px solid ${C.b0}`,borderRadius:12,padding:'18px 20px'}}>
            <div style={{fontSize:12,fontWeight:700,color:C.t2,marginBottom:16,textTransform:'uppercase',letterSpacing:.7}}>📊 Severity Dağılımı</div>
            {Object.keys(SEV).map(sev=>{
              const count=bugs.filter(b=>b.severity===sev).length
              const pct=total?Math.round(count/total*100):0
              const sc=SEV[sev]
              return(
                <div key={sev} style={{marginBottom:10}}>
                  <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:4}}>
                    <div style={{display:'flex',alignItems:'center',gap:7}}>
                      <div style={{width:7,height:7,borderRadius:'50%',background:sc.fg,boxShadow:`0 0 5px ${sc.fg}88`}}/>
                      <span style={{fontSize:12,color:sc.fg,fontWeight:600}}>{sev}</span>
                    </div>
                    <div style={{display:'flex',alignItems:'center',gap:7}}>
                      <span style={{fontSize:11,color:C.t3}}>{count} bug</span>
                      <Tag color={sc.fg}>{pct}%</Tag>
                    </div>
                  </div>
                  <div style={{height:8,background:'rgba(255,255,255,0.04)',borderRadius:4,overflow:'hidden'}}>
                    <div style={{height:'100%',width:`${pct}%`,background:`linear-gradient(90deg,${sc.fg},${sc.fg}99)`,borderRadius:4,transition:'width .8s ease',boxShadow:`0 0 6px ${sc.fg}55`}}/>
                  </div>
                </div>
              )
            })}
            {total===0&&<div style={{textAlign:'center',padding:'16px 0',color:C.t3,fontSize:12}}>Bug eklendikçe burada görünür</div>}
          </div>

          {/* Version trend */}
          <div style={{background:C.bg2,border:`1px solid ${C.b0}`,borderRadius:12,padding:'18px 20px'}}>
            <div style={{fontSize:12,fontWeight:700,color:C.t2,marginBottom:16,textTransform:'uppercase',letterSpacing:.7}}>📈 Versiyon Trendi</div>
            {versions.length===0?(
              <div style={{textAlign:'center',padding:'20px 0',color:C.t3,fontSize:12}}>Versiyon bazlı bug yok</div>
            ):versions.map(([ver,count],i)=>{
              const prev=i>0?versions[i-1][1]:null
              const tr=prev===null?'':count>prev?'↑':count<prev?'↓':'→'
              const trC=tr==='↑'?C.red:tr==='↓'?C.green:C.yellow
              return(
                <div key={ver} style={{marginBottom:10}}>
                  <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:4}}>
                    <Tag color={C.blue}>v{ver}</Tag>
                    <span style={{fontSize:12,color:C.t3,display:'flex',alignItems:'center',gap:5}}>
                      <span style={{fontWeight:700,color:C.t1}}>{count}</span> bug
                      {tr&&<span style={{color:trC,fontWeight:800,fontSize:14}}>{tr}</span>}
                    </span>
                  </div>
                  <div style={{height:8,background:'rgba(255,255,255,0.04)',borderRadius:4,overflow:'hidden'}}>
                    <div style={{height:'100%',width:`${Math.round(count/maxVer*100)}%`,background:`linear-gradient(90deg,${C.blue},${C.purple})`,borderRadius:4,boxShadow:`0 0 8px ${C.blue}44`}}/>
                  </div>
                </div>
              )
            })}
          </div>
        </div>

        {/* Top tools */}
        <div style={{background:C.bg2,border:`1px solid ${C.b0}`,borderRadius:12,padding:'18px 20px'}}>
          <div style={{fontSize:12,fontWeight:700,color:C.t2,marginBottom:14,textTransform:'uppercase',letterSpacing:.7}}>⭐ En Çok Kullanılan Araçlar</div>
          {topTools.length===0?(
            <div style={{color:C.t3,fontSize:12,textAlign:'center',padding:'12px 0'}}>Araç çalıştırıldıkça burada görünür</div>
          ):topTools.map(([id,count])=>{
            const t=allTools.find(x=>x.id===id); if(!t) return null
            const cl=PALETTE[t.color]||PALETTE.blue
            return(
              <div key={id} style={{display:'flex',alignItems:'center',gap:12,marginBottom:10,padding:'8px 10px',borderRadius:8,background:'rgba(255,255,255,0.01)'}}>
                <span style={{fontSize:17}}>{t.icon}</span>
                <div style={{flex:1}}>
                  <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:4}}>
                    <span style={{fontSize:12,color:C.t1,fontWeight:600}}>{t.label}</span>
                    <Tag color={cl.text}>{count}×</Tag>
                  </div>
                  <div style={{height:5,background:'rgba(255,255,255,0.04)',borderRadius:3,overflow:'hidden'}}>
                    <div style={{height:'100%',width:`${Math.round(count/maxUse*100)}%`,background:cl.text,borderRadius:3}}/>
                  </div>
                </div>
              </div>
            )
          })}
        </div>
      </div>
    </div>
  )
}

// ─── TOOL MANAGER ─────────────────────────────────────────────────────────────

// ─── TOOL TYPE ACCORDION ──────────────────────────────────────────────────────
function ToolTypeAccordion({form, setForm}){
  const groups = Object.entries(TOOL_TYPES).reduce((acc,[key,tt])=>{
    const g = tt.group||'Diğer'
    if(!acc[g]) acc[g] = []
    acc[g].push([key,tt])
    return acc
  },{})

  const [openGroups, setOpenGroups] = useState(()=>{
    // Seçili toolType'ın grubunu aç, diğerlerini kapat
    const init = {}
    Object.entries(groups).forEach(([g,items])=>{
      init[g] = items.some(([k])=>k===form.toolType)
    })
    // En az bir grup açık olsun
    if(!Object.values(init).some(Boolean)) init[Object.keys(groups)[0]] = true
    return init
  })

  const toggleGroup = (g) => setOpenGroups(o=>({...o,[g]:!o[g]}))

  return(
    <div style={{background:'rgba(255,255,255,0.02)',border:`1px solid ${C.b1}`,borderRadius:10,marginBottom:16,overflow:'hidden'}}>
      {/* Başlık */}
      <div style={{display:'flex',alignItems:'center',gap:8,padding:'10px 14px',borderBottom:`1px solid ${C.b0}`}}>
        <span style={{fontSize:11,fontWeight:700,color:C.t2,textTransform:'uppercase',letterSpacing:.7,flex:1}}>🎯 Araç Tipi — Ne üretsin?</span>
        {form.toolType&&(
          <span style={{fontSize:10,padding:'2px 8px',borderRadius:5,
            background:`${TOOL_TYPES[form.toolType]?.color}18`,
            color:TOOL_TYPES[form.toolType]?.color,
            border:`1px solid ${TOOL_TYPES[form.toolType]?.color}40`,fontWeight:600}}>
            {TOOL_TYPES[form.toolType]?.icon} {TOOL_TYPES[form.toolType]?.label}
          </span>
        )}
      </div>

      {/* Grup accordion'ları */}
      {Object.entries(groups).map(([group, items])=>{
        const isOpen = openGroups[group]
        const hasSelected = items.some(([k])=>k===form.toolType)
        return(
          <div key={group} style={{borderBottom:`1px solid ${C.b0}`}}>
            {/* Grup başlığı */}
            <div onClick={()=>toggleGroup(group)}
              style={{display:'flex',alignItems:'center',gap:6,padding:'8px 14px',cursor:'pointer',
                background:isOpen?'rgba(255,255,255,0.02)':'transparent',
                transition:'background .15s',userSelect:'none'}}
              onMouseEnter={e=>e.currentTarget.style.background='rgba(255,255,255,0.03)'}
              onMouseLeave={e=>e.currentTarget.style.background=isOpen?'rgba(255,255,255,0.02)':'transparent'}>
              <span style={{fontSize:10,color:C.t3,transition:'transform .2s',display:'inline-block',
                transform:isOpen?'rotate(90deg)':'rotate(0deg)',flexShrink:0}}>▶</span>
              <span style={{fontSize:11,fontWeight:700,color:isOpen?C.t2:C.t3,textTransform:'uppercase',letterSpacing:.7,flex:1}}>{group}</span>
              {hasSelected&&<span style={{fontSize:10,color:C.blue,fontWeight:700}}>✓ seçili</span>}
              <span style={{fontSize:10,color:C.t3}}>{items.length}</span>
            </div>
            {/* Grup içeriği */}
            {isOpen&&(
              <div style={{padding:'8px 12px 10px',display:'flex',flexWrap:'wrap',gap:5}}>
                {items.map(([key,tt])=>(
                  <div key={key} onClick={()=>setForm(f=>({...f,toolType:key}))}
                    style={{display:'flex',alignItems:'center',gap:6,padding:'6px 10px',borderRadius:7,cursor:'pointer',
                      border:`1px solid ${form.toolType===key?tt.color+'70':C.b0}`,
                      background:form.toolType===key?`${tt.color}18`:'rgba(255,255,255,0.01)',
                      transition:'all .12s'}}>
                    <span style={{fontSize:14}}>{tt.icon}</span>
                    <div>
                      <div style={{fontSize:11,fontWeight:700,color:form.toolType===key?tt.color:C.t1,whiteSpace:'nowrap'}}>{tt.label}</div>
                      <div style={{fontSize:9,color:C.t3,lineHeight:1.3}}>{tt.outputLabel||'Metin'}</div>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )
      })}

      {/* Seçili tip bilgisi */}
      {form.toolType&&form.toolType!=='chat'&&(
        <div style={{padding:'8px 14px',background:`${TOOL_TYPES[form.toolType]?.color}08`,fontSize:11,color:TOOL_TYPES[form.toolType]?.color}}>
          💡 Çıktı: <strong>{TOOL_TYPES[form.toolType]?.outputLabel}</strong> — çalıştırıldığında indirilebilir dosya üretecek
        </div>
      )}
    </div>
  )
}

function ToolManagerTab({customTools, setCustomTools}){
  const [form,setForm]=useState({label:'',desc:'',cat:'Analiz',icon:'🔧',color:'blue',prompt:'',toolType:'chat'})
  const [editId,setEditId]=useState(null)
  const [aiInput,sAI]=useState('')
  const [aiLoading,sAL]=useState(false)
  const [saved,sSaved]=useState(false)

  const startEdit=(t)=>{
    setForm({label:t.label,desc:t.desc||'',cat:t.cat,icon:t.icon,color:t.color,prompt:t.prompt||'',toolType:t.toolType||'chat'})
    setEditId(t.id)
    window.scrollTo(0,0)
  }

  const cancelEdit=()=>{
    setEditId(null)
    setForm({label:'',desc:'',cat:'Analiz',icon:'🔧',color:'blue',prompt:'',toolType:'chat'})
  }
  const icons=['🔧','📊','🔍','🧪','⚙️','🛡','📡','🗄','📧','🧩','📦','🔐','🧬','📱','💡','🎯','🔬','🚀','⚡','🌐','🔑','📌','🔔','🗃','💎','🧲','🗝','🏷','📐','🔩','🖥️','🐍','🎭','🦊','🐛','🦋','🔮','🧠','🏗','🔄','📋','✅','❌','⚠️','🚨','🔴','🟠','🟡','🟢','🔵','🟣','⬛','📝','🗒','📂','🗂','💾','💿','🖨','⌨️','🖱','📺','📷','🎥','🔊','📻','🛰','🌍','🔗','⛓','🧱','🏛','🔭','⚗️','🧫','🧰','🪛','🔨','⚒️','🛠','🪜','🧯','🚦','🚧','🏁','🎮','🕹','🎲','🎯','🏆','🥇','🎖','📈','📉','💹','💰','💳','🏦','🏢','🏭']

  const addTool=()=>{
    if(!form.label.trim()||!form.prompt.trim()) return
    let upd
    if(editId){
      upd = customTools.map(t=>t.id===editId?{...form,id:editId}:t)
      setEditId(null)
    } else {
      upd = [...customTools, {...form, id:'custom_'+Date.now()}]
    }
    try { localStorage.setItem('qa_custom_tools', JSON.stringify(upd)) } catch(e) { alert('Kayıt hatası: '+e.message); return }
    setCustomTools(upd)
    setForm({label:'',desc:'',cat:'Analiz',icon:'🔧',color:'blue',prompt:'',toolType:'chat'})
    sSaved(true); setTimeout(()=>sSaved(false),2500)
  }

  const genAI=async()=>{
    if(!aiInput.trim()||aiLoading) return
    sAL(true)
    try{
      // Kullanıcının seçtiği toolType'ı koru — AI ezmesin
      const currentType = form.toolType || 'chat'
      const typeInfo = TOOL_TYPES[currentType]
      const prompt = `QA Studio için araç tasarla.
Araç tipi: ${currentType} (${typeInfo?.label || currentType})
Kullanıcı isteği: "${aiInput}"

Sadece aşağıdaki JSON formatında yanıt ver, başka hiçbir şey yazma:
{"label":"kısa araç adı","desc":"ne yapar (kısa)","cat":"Analiz","icon":"🔧","color":"blue","prompt":"Bu araç çalıştığında AI ya verilecek Türkçe detaylı prompt. toolType ${currentType} olduğu için ${currentType === 'chat' ? 'analiz ve metin çıktısı' : typeInfo?.label + ' kodu'} üretmeye yönelik olsun."}`
      const r = await callAI([{role:'user', content: prompt}])
      // JSON'u temizle ve parse et
      const clean = r.replace(/```json/g,'').replace(/```/g,'').trim()
      const parsed = JSON.parse(clean)
      // toolType'ı kullanıcının seçiminden al, AI'dan alma
      setForm(f=>({...f, ...parsed, toolType: currentType}))
      sAI('')
    }catch(e){
      alert('AI yanıtı işlenemedi: ' + e.message + '\nTekrar deneyin.')
    }
    finally{sAL(false)}
  }

  return(
    <div style={{display:'flex',height:'100%',background:C.bg0,overflow:'hidden'}}>
      {/* Form panel */}
      <div style={{flex:1,overflowY:'auto',padding:24,borderRight:`1px solid ${C.b0}`}}>
        <div style={{maxWidth:640}}>
          <div style={{fontSize:15,fontWeight:800,color:C.t0,marginBottom:20}}>{editId?'✏️ Aracı Düzenle':'Yeni Araç Oluştur'}</div>

          {/* Araç Tipi Seçimi — Accordion */}
          <ToolTypeAccordion form={form} setForm={setForm}/>

          {/* AI generator */}
          <div style={{background:'rgba(139,92,246,0.06)',border:`1px solid rgba(139,92,246,0.2)`,borderRadius:10,padding:'14px 16px',marginBottom:20}}>
            <div style={{fontSize:11,fontWeight:700,color:C.purple,textTransform:'uppercase',letterSpacing:.7,marginBottom:10}}>🤖 AI ile Otomatik Oluştur</div>
            <div style={{display:'flex',gap:8}}>
              <TextInput value={aiInput} onChange={e=>sAI(e.target.value)} onKeyDown={e=>e.key==='Enter'&&genAI()} placeholder="Ne tür araç istiyorsun? Örn: GraphQL şema testi"/>
              <BtnAccent onClick={genAI} disabled={!aiInput.trim()} loading={aiLoading} icon="✨">Oluştur</BtnAccent>
            </div>
            <div style={{fontSize:11,color:C.t3,marginTop:8}}>AI formu otomatik dolduracak, istediğin gibi düzenleyip kaydedebilirsin</div>
          </div>

          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:14,marginBottom:14}}>
            <Field label="Araç Adı *"><TextInput value={form.label} onChange={e=>setForm(f=>({...f,label:e.target.value}))} placeholder="Örn: GraphQL Testi"/></Field>
            <Field label="Açıklama"><TextInput value={form.desc} onChange={e=>setForm(f=>({...f,desc:e.target.value}))} placeholder="Ne yapar?"/></Field>
          </div>

          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:14,marginBottom:14}}>
            <Field label="Kategori">
              <Sel value={form.cat} onChange={e=>setForm(f=>({...f,cat:e.target.value}))}>
                {CATEGORIES.map(c=><option key={c}>{c}</option>)}
              </Sel>
            </Field>
            <Field label="Renk">
              <div style={{display:'flex',gap:6,flexWrap:'wrap',paddingTop:8}}>
                {Object.keys(PALETTE).map(c=><div key={c} onClick={()=>setForm(f=>({...f,color:c}))} title={c} style={{width:24,height:24,borderRadius:6,background:PALETTE[c].text,cursor:'pointer',border:form.color===c?`2px solid #fff`:'2px solid transparent',transition:'all .15s',transform:form.color===c?'scale(1.2)':'scale(1)'}}/>)}
              </div>
            </Field>
          </div>

          <Field label="İkon">
            <div style={{display:'flex',gap:5,flexWrap:'wrap',paddingTop:6}}>
              {icons.map(ic=><div key={ic} onClick={()=>setForm(f=>({...f,icon:ic}))} style={{width:34,height:34,borderRadius:7,display:'flex',alignItems:'center',justifyContent:'center',fontSize:17,cursor:'pointer',background:form.icon===ic?'rgba(59,158,237,0.15)':'rgba(255,255,255,0.03)',border:`1px solid ${form.icon===ic?C.blue:C.b0}`,transition:'all .15s'}}>{ic}</div>)}
            </div>
          </Field>

          <div style={{marginTop:14}}>
            <Field label="AI Prompt *">
              <TextArea value={form.prompt} onChange={e=>setForm(f=>({...f,prompt:e.target.value}))} rows={5} placeholder="Bu araç çalıştığında AI'ya ne söylensin? Girilen konuya göre ne üretmeli?"/>
            </Field>
          </div>

          {form.label&&(
            <div style={{marginTop:14,marginBottom:14}}>
              <div style={{fontSize:11,fontWeight:600,color:C.t3,textTransform:'uppercase',letterSpacing:.7,marginBottom:8}}>Önizleme</div>
              <ToolChip tool={form} selected={false} onClick={()=>{}}/>
            </div>
          )}

          <div style={{display:'flex',gap:8,marginTop:18,paddingTop:16,borderTop:`1px solid ${C.b0}`}}>
            <BtnPrimary size="lg" onClick={addTool} disabled={!form.label.trim()||!form.prompt.trim()}>{saved?'✓ Kaydedildi!':editId?'💾 Güncelle':'+ Aracı Kaydet'}</BtnPrimary>
            {editId
              ? <BtnGhost size="lg" onClick={cancelEdit}>✕ İptal</BtnGhost>
              : <BtnGhost size="lg" onClick={()=>setForm({label:'',desc:'',cat:'Analiz',icon:'🔧',color:'blue',prompt:''})}>Formu Temizle</BtnGhost>
            }
          </div>
        </div>
      </div>

      {/* Custom tools list */}
      <div style={{width:340,flexShrink:0,overflowY:'auto',padding:20,background:C.bg1,borderLeft:`1px solid ${C.b0}`}}>
        <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:16}}>
          <span style={{fontSize:13,fontWeight:700,color:C.t0}}>Özel Araçlarım</span>
          <Tag color={C.blue}>{customTools.length}</Tag>
        </div>
        {customTools.length===0?(
          <div style={{textAlign:'center',padding:'40px 0',color:C.t3}}>
            <div style={{fontSize:34,marginBottom:10,opacity:.4}}>🧰</div>
            <div style={{fontSize:12}}>Henüz özel araç eklenmedi</div>
          </div>
        ):(
          <div>
            {customTools.map(t=>(
              <CustomToolRow key={t.id} tool={t} isEditing={editId===t.id} onEdit={()=>startEdit(t)} onDelete={()=>askConfirm(`"${t.label}" aracını kaldırmak istediğinize emin misiniz?`, ()=>{const u=customTools.filter(x=>x.id!==t.id);setCustomTools(u);LS.set('qa_custom_tools',u);if(editId===t.id)cancelEdit();setConfirm(null)})} toggle={()=>{}} selectedTools={[]}/>
            ))}
          </div>
        )}
      </div>
    </div>
  )
}


// ─── TEST SÜREÇLERİ TAB ───────────────────────────────────────────────────────
function ProcessTab(){
  const defaultProcesses = [
    { id:1, name:'Gereksinim Analizi',    responsible:'Sistem Müh. / BA',  next:'Test Planlama',         desc:'Gereksinimlerin test edilebilirlik açısından incelenmesi, belirsizliklerin giderilmesi' },
    { id:2, name:'Test Planlama',         responsible:'Test Yöneticisi',   next:'Test Tasarımı',         desc:'Test stratejisi, kapsam, kaynaklar, takvim ve risk planının oluşturulması' },
    { id:3, name:'Test Tasarımı',         responsible:'Kıdemli Test Müh.', next:'Test Geliştirme',       desc:'Test case, senaryo ve veri tasarımı; traceability matrisinin hazırlanması' },
    { id:4, name:'Test Geliştirme',       responsible:'Test Müh.',         next:'Test Ortamı Kurulum',   desc:'Test caselerinin dokümante edilmesi, otomasyon scriptlerinin yazılması' },
    { id:5, name:'Test Ortamı Kurulum',   responsible:'DevOps / Test Müh.',next:'Test Yürütme',          desc:'Test ortamının hazırlanması, araçların kurulması, veri yüklenmesi' },
    { id:6, name:'Test Yürütme',          responsible:'Test Müh.',         next:'Hata Yönetimi',         desc:'Test caselerinin koşturulması, sonuçların kaydedilmesi' },
    { id:7, name:'Hata Yönetimi',         responsible:'Test Müh. / Geliştirici', next:'Regresyon Testi', desc:'Hataların raporlanması, önceliklendirilmesi, takibi ve kapanması' },
    { id:8, name:'Regresyon Testi',       responsible:'Test Müh.',         next:'Test Raporu',           desc:'Düzeltmelerin doğrulanması, mevcut fonksiyonların bozulmadığının kontrolü' },
    { id:9, name:'Test Raporu',           responsible:'Test Yöneticisi',   next:'Sürüm Onayı',           desc:'Test özeti, metrikler, kapsam ve tavsiye raporunun hazırlanması' },
    { id:10,name:'Sürüm Onayı',           responsible:'Test Yöneticisi / PM', next:'Kapatma',            desc:'Çıkış kriterlerinin değerlendirilmesi, sürüm için onay kararı' },
  ]
  const [processes, setProcesses] = useState(()=>LS.get('qa_processes', defaultProcesses))
  const [selected, setSelected] = useState(null)
  const [editMode, setEditMode] = useState(false)
  const [editItem, setEditItem] = useState(null)
  const [action, setAction] = useState('')
  const [actionResult, setActionResult] = useState('')
  const [loading, setLoading] = useState(false)

  const saveProc = (p) => { setProcesses(p); LS.set('qa_processes', p) }

  const startEdit = (item) => { setEditItem({...item}); setEditMode(true) }
  const saveEdit = () => {
    if(!editItem) return
    const isNew = !processes.find(p=>p.id===editItem.id)
    saveProc(isNew ? [...processes, editItem] : processes.map(p=>p.id===editItem.id?editItem:p))
    setEditMode(false); setEditItem(null)
  }

  const askAI = async() => {
    if(!selected||!action.trim()) return
    setLoading(true); setActionResult('')
    try{
      const proc = processes.find(p=>p.id===selected)
      const r = await callAI([{role:'user',content:`Test süreci: "${proc.name}"
Sorumlu: ${proc.responsible}
Açıklama: ${proc.desc}

Kullanıcı sorusu/isteği: ${action}

Türkçe, somut ve uygulanabilir yanıt ver.`}])
      setActionResult(r)
    }catch(e){setActionResult('Hata: '+e.message)}
    finally{setLoading(false)}
  }

  return(
    <div style={{display:'flex',height:'100%',overflow:'hidden'}}>
      {/* Süreç listesi */}
      <div style={{width:300,flexShrink:0,borderRight:`1px solid ${C.b0}`,overflowY:'auto',background:C.bg1}}>
        <div style={{padding:'14px 16px',borderBottom:`1px solid ${C.b0}`,display:'flex',alignItems:'center',gap:8}}>
          <span style={{fontSize:13,fontWeight:700,color:C.t0,flex:1}}>🔀 Test Süreçleri</span>
          <BtnGhost size="sm" onClick={()=>{setEditItem({id:Date.now(),name:'',responsible:'',next:'',desc:''});setEditMode(true)}}>+ Ekle</BtnGhost>
        </div>
        {processes.map((p,i)=>(
          <div key={p.id} onClick={()=>setSelected(p.id)}
            style={{padding:'10px 16px',borderBottom:`1px solid ${C.b0}`,cursor:'pointer',
              background:selected===p.id?'rgba(59,158,237,0.08)':'transparent',
              borderLeft:`3px solid ${selected===p.id?C.blue:'transparent'}`,
              transition:'all .15s'}}>
            <div style={{display:'flex',alignItems:'center',gap:6}}>
              <span style={{fontSize:11,fontWeight:700,color:C.t3,flexShrink:0,width:20}}>{i+1}.</span>
              <span style={{fontSize:12,fontWeight:600,color:selected===p.id?C.blue:C.t1,flex:1}}>{p.name}</span>
              <button onClick={e=>{e.stopPropagation();startEdit(p)}} style={{fontSize:10,padding:'2px 6px',borderRadius:4,border:`1px solid ${C.b0}`,background:'transparent',color:C.t3,cursor:'pointer',fontFamily:'inherit'}}>✏️</button>
            </div>
            <div style={{fontSize:11,color:C.t3,marginTop:3,paddingLeft:26}}>{p.responsible}</div>
          </div>
        ))}
      </div>

      {/* Süreç detayı + AI */}
      <div style={{flex:1,overflowY:'auto',padding:24}}>
        {!selected&&!editMode&&(
          <div style={{textAlign:'center',padding:'60px 0',color:C.t3}}>
            <div style={{fontSize:40,marginBottom:12,opacity:.3}}>🔀</div>
            <div style={{fontSize:14,fontWeight:700,color:C.t2,marginBottom:6}}>Bir süreç adımı seç</div>
            <div style={{fontSize:12}}>Sorumluluğu ve sonraki adımı görüntüle, AI ile derinleştir</div>
          </div>
        )}
        {selected&&!editMode&&(()=>{
          const p = processes.find(x=>x.id===selected)
          if(!p) return null
          return(
            <div style={{maxWidth:640}}>
              <div style={{fontSize:20,fontWeight:800,color:C.t0,marginBottom:6}}>{p.name}</div>
              <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:12,marginBottom:20}}>
                <div style={{padding:'12px 16px',borderRadius:9,background:'rgba(59,158,237,0.06)',border:`1px solid rgba(59,158,237,0.15)`}}>
                  <div style={{fontSize:10,color:C.blue,fontWeight:700,textTransform:'uppercase',letterSpacing:.7,marginBottom:4}}>Sorumlu</div>
                  <div style={{fontSize:13,color:C.t0,fontWeight:600}}>{p.responsible}</div>
                </div>
                <div style={{padding:'12px 16px',borderRadius:9,background:'rgba(16,185,129,0.06)',border:`1px solid rgba(16,185,129,0.15)`}}>
                  <div style={{fontSize:10,color:C.green,fontWeight:700,textTransform:'uppercase',letterSpacing:.7,marginBottom:4}}>Sonraki Adım</div>
                  <div style={{fontSize:13,color:C.t0,fontWeight:600}}>{p.next}</div>
                </div>
              </div>
              <div style={{padding:'14px 16px',borderRadius:9,background:'rgba(255,255,255,0.02)',border:`1px solid ${C.b0}`,marginBottom:20}}>
                <div style={{fontSize:11,color:C.t3,fontWeight:700,textTransform:'uppercase',letterSpacing:.7,marginBottom:6}}>Açıklama</div>
                <div style={{fontSize:13,color:C.t1,lineHeight:1.7}}>{p.desc}</div>
              </div>
              <div style={{background:'rgba(139,92,246,0.06)',border:`1px solid rgba(139,92,246,0.2)`,borderRadius:10,padding:'14px 16px'}}>
                <div style={{fontSize:11,fontWeight:700,color:C.purple,textTransform:'uppercase',letterSpacing:.7,marginBottom:10}}>🤖 Bu Süreçte Ne Yapmak İstiyorsun?</div>
                <div style={{display:'flex',gap:8,marginBottom:8}}>
                  <TextInput value={action} onChange={e=>setAction(e.target.value)} onKeyDown={e=>e.key==='Enter'&&askAI()} placeholder={`"${p.name}" için ne yapmamı istersin?`}/>
                  <BtnAccent onClick={askAI} loading={loading} disabled={!action.trim()}>Sor</BtnAccent>
                </div>
                {actionResult&&<pre style={{whiteSpace:'pre-wrap',fontSize:12,color:C.t1,lineHeight:1.7,marginTop:10,padding:'10px 12px',background:'rgba(0,0,0,0.2)',borderRadius:7}}>{actionResult}</pre>}
              </div>
            </div>
          )
        })()}
        {editMode&&editItem&&(
          <div style={{maxWidth:500}}>
            <div style={{fontSize:14,fontWeight:700,color:C.t0,marginBottom:16}}>{editItem.id&&processes.find(p=>p.id===editItem.id)?'Adımı Düzenle':'Yeni Adım Ekle'}</div>
            <div style={{display:'flex',flexDirection:'column',gap:12}}>
              <Field label="Adım Adı"><TextInput value={editItem.name} onChange={e=>setEditItem(x=>({...x,name:e.target.value}))} placeholder="Test Planlama"/></Field>
              <Field label="Sorumlu"><TextInput value={editItem.responsible} onChange={e=>setEditItem(x=>({...x,responsible:e.target.value}))} placeholder="Test Yöneticisi"/></Field>
              <Field label="Sonraki Adım"><TextInput value={editItem.next} onChange={e=>setEditItem(x=>({...x,next:e.target.value}))} placeholder="Test Tasarımı"/></Field>
              <Field label="Açıklama"><TextArea value={editItem.desc} onChange={e=>setEditItem(x=>({...x,desc:e.target.value}))} rows={3} placeholder="Bu adımda ne yapılır?"/></Field>
            </div>
            <div style={{display:'flex',gap:8,marginTop:16}}>
              <BtnPrimary onClick={saveEdit} disabled={!editItem.name.trim()}>💾 Kaydet</BtnPrimary>
              <BtnGhost onClick={()=>{setEditMode(false);setEditItem(null)}}>İptal</BtnGhost>
              {processes.find(p=>p.id===editItem.id)&&<BtnDanger onClick={()=>{saveProc(processes.filter(p=>p.id!==editItem.id));setEditMode(false);setEditItem(null);setSelected(null)}}>Sil</BtnDanger>}
            </div>
          </div>
        )}
      </div>
    </div>
  )
}

// ─── SAVUNMA SANAYİİ TAB ──────────────────────────────────────────────────────
function DefenseTab(){
  const [activeSection, setActiveSection] = useState('overview')
  const [input, setInput] = useState('')
  const [selectedTools, setSelectedTools] = useState([])
  const [result, setResult] = useState('')
  const [loading, setLoading] = useState(false)
  const [exportFmt, setExportFmt] = useState('txt')

  const sections = [
    { id:'overview', label:'📋 Genel Bakış',        color:C.blue   },
    { id:'vv',       label:'🔬 V&V Süreci',          color:C.teal   },
    { id:'milstd',   label:'🛡️ MIL-STD',            color:C.red    },
    { id:'icd',      label:'📡 ICD/Arayüz',          color:C.purple },
    { id:'env',      label:'🌡️ Çevresel Test',      color:C.orange },
    { id:'cyber',    label:'🔐 Siber Güvenlik',       color:C.red    },
    { id:'fat_sat',  label:'🏭 FAT/SAT',              color:C.green  },
    { id:'ai_assist',label:'🤖 Serbest AI',           color:C.purple },
  ]

  // Her bölüm için araçlar
  const SECTION_TOOLS = {
    overview: [
      { id:'ov1', icon:'📊', label:'V&V Analiz Raporu',      desc:'Sistem → doğrulama ve geçerleme analizi üret' },
      { id:'ov2', icon:'📋', label:'Test Strateji Belgesi',   desc:'Proje → savunma odaklı test stratejisi' },
      { id:'ov3', icon:'🔗', label:'Traceability Matrisi',    desc:'Gereksinim → test case izlenebilirlik matrisi' },
      { id:'ov4', icon:'⚠️', label:'Risk Analizi',           desc:'Sistem → risk sınıflandırma ve önceliklendirme' },
    ],
    vv: [
      { id:'vv1', icon:'🔬', label:'V&V Test Planı',          desc:'Sistem → MIL-STD-498 uyumlu V&V test planı' },
      { id:'vv2', icon:'📝', label:'Test Prosedürü (TPr)',     desc:'Senaryo → adım adım test prosedürü belgesi' },
      { id:'vv3', icon:'📑', label:'Test Raporu (TR)',         desc:'Sonuçlar → resmi test raporu formatı' },
      { id:'vv4', icon:'🏗', label:'V Modeli Senaryo',         desc:'Proje seviyesi → V modeli test senaryoları' },
      { id:'vv5', icon:'✅', label:'Giriş/Çıkış Kriterleri',  desc:'Test fazı → geçiş kriterleri listesi' },
    ],
    milstd: [
      { id:'ms1', icon:'🛡️', label:'MIL-STD-498 Kontrol',    desc:'Yazılım → MIL-STD-498 uyumluluk kontrol listesi' },
      { id:'ms2', icon:'✈️', label:'DO-178C Analizi',         desc:'Havacılık yazılımı → DO-178C seviye analizi' },
      { id:'ms3', icon:'🔒', label:'MIL-STD-882 Güvenlik',    desc:'Sistem → güvenlik kritiklik analizi (SSPP)' },
      { id:'ms4', icon:'📡', label:'MIL-STD-1553 Test',       desc:'Veri yolu → 1553B protokol test senaryoları' },
      { id:'ms5', icon:'⚡', label:'MIL-STD-461 EMC/EMI',     desc:'Sistem → EMI/EMC test gereksinimleri' },
    ],
    icd: [
      { id:'ic1', icon:'📡', label:'ICD Test Senaryosu',      desc:'ICD dokümanı → arayüz test senaryoları' },
      { id:'ic2', icon:'📨', label:'Mesaj Format Testi',       desc:'Protokol → mesaj doğrulama test adımları' },
      { id:'ic3', icon:'⏱️', label:'Zamanlama Analizi',       desc:'Arayüz → timing ve periyot test senaryoları' },
      { id:'ic4', icon:'🔧', label:'Hata Enjeksiyon Testi',   desc:'Arayüz → hata durumu test senaryoları' },
      { id:'ic5', icon:'🔗', label:'IRS Gereksinim Doğrulama',desc:'IRS → gereksinim doğrulama test matrisi' },
    ],
    env: [
      { id:'en1', icon:'🌡️', label:'Termal Test Planı',      desc:'Sistem → MIL-STD-810 termal test prosedürü' },
      { id:'en2', icon:'📳', label:'Titreşim Test Planı',     desc:'Sistem → rastgele/sinüzoidal titreşim testi' },
      { id:'en3', icon:'💥', label:'Şok Test Planı',          desc:'Sistem → yarı-sinüs ve darbe şok test planı' },
      { id:'en4', icon:'💧', label:'Nem/Su Geçirmezlik',      desc:'Sistem → IP67/68 ve nem dayanım test planı' },
      { id:'en5', icon:'🌬️', label:'Kum/Toz Testi',          desc:'Sistem → MIL-STD-810 Method 510 test prosedürü' },
      { id:'en6', icon:'🔌', label:'Yazılım Çevresel Testi',  desc:'Yazılım → termal değişim, güç kesintisi test senaryoları' },
    ],
    cyber: [
      { id:'cy1', icon:'🔐', label:'Sızma Testi Planı',       desc:'Sistem → siber güvenlik sızma test planı' },
      { id:'cy2', icon:'🛡️', label:'TEMPEST Test Senaryosu',  desc:'Donanım → elektromanyetik emanasyon testleri' },
      { id:'cy3', icon:'🔑', label:'Kimlik Doğrulama Testi',  desc:'Sistem → auth/authz test senaryoları' },
      { id:'cy4', icon:'🌐', label:'Ağ Güvenlik Testi',       desc:'Ağ → port tarama, firewall test senaryoları' },
      { id:'cy5', icon:'💉', label:'Zafiyet Analizi',          desc:'Sistem → CVE ve zafiyet tarama planı' },
    ],
    fat_sat: [
      { id:'fs1', icon:'🏭', label:'FAT Test Prosedürü',      desc:'Sistem → fabrika kabul test prosedürü belgesi' },
      { id:'fs2', icon:'🏗', label:'SAT Test Prosedürü',      desc:'Sistem → saha kabul test prosedürü belgesi' },
      { id:'fs3', icon:'✅', label:'Kabul Kriterleri',         desc:'Test fazı → FAT/SAT geçiş kriterleri listesi' },
      { id:'fs4', icon:'📋', label:'FAT Kontrol Listesi',      desc:'FAT hazırlık → eksiksiz kontrol listesi' },
      { id:'fs5', icon:'📊', label:'FAT Sonuç Raporu',         desc:'FAT sonuçları → resmi kabul test raporu' },
    ],
  }

  const currentTools = SECTION_TOOLS[activeSection] || []
  const toggleTool = (id) => setSelectedTools(t => t.includes(id) ? t.filter(x=>x!==id) : [...t,id])

  const PROMPTS = {
    ov1: 'Savunma sistemi V&V analiz raporu hazırla. Doğrulama ve geçerleme faaliyetlerini, yöntemlerini ve başarı kriterlerini detaylandır.',
    ov2: 'MIL-STD ve IEEE 829 uyumlu savunma sistemi test stratejisi belgesi hazırla. Kapsam, yaklaşım, araçlar, takvim ve sorumlulukları içersin.',
    ov3: 'Gereksinim-test case izlenebilirlik matrisi oluştur. Her gereksinim için test yöntemini, ilgili test caselerini ve doğrulama durumunu tablola.',
    ov4: 'Savunma sistemi risk analizi yap. Tehdit tanımlama, olasılık-şiddet matrisi, risk sınıflandırma ve azaltma önerilerini içersin.',
    vv1: 'MIL-STD-498 uyumlu V&V test planı hazırla. Birim, entegrasyon, sistem ve kabul test seviyelerini, giriş/çıkış kriterlerini içersin.',
    vv2: 'Adım adım test prosedürü (TPr) belgesi oluştur. Her adım için ön koşul, işlem, beklenen sonuç ve geçiş kriterini belirt.',
    vv3: 'Resmi test raporu (TR) formatı oluştur. Test özeti, yürütülen testler, sonuçlar, anomaliler ve tavsiyeler bölümlerini içersin.',
    vv4: 'V modeli her seviyesi için test senaryoları üret. Birimden sisteme her seviyede test kapsamı, yöntemi ve dokümantasyon gereksinimlerini belirt.',
    vv5: 'Test fazı giriş ve çıkış kriterleri listesi oluştur. Her kriter için ölçülebilir kabul şartlarını belirt.',
    ms1: 'MIL-STD-498 uyumluluk kontrol listesi oluştur. Her gereksinim maddesi için doğrulama yöntemi ve durum kolonları içersin.',
    ms2: 'DO-178C yazılım seviye analizi yap. DAL seviyeleri, uçuş kritiklik fonksiyonları ve sertifikasyon gereksinimlerini analiz et.',
    ms3: 'MIL-STD-882 sistem güvenlik program planı (SSPP) oluştur. Tehlike tanımlama, risk azaltma ve takip süreçlerini içersin.',
    ms4: 'MIL-STD-1553B veri yolu protokol test senaryoları üret. BC/RT/BM iletişimi, hata koşulları ve zamanlama testlerini içersin.',
    ms5: 'MIL-STD-461G EMI/EMC test gereksinimleri ve test prosedürleri listesi oluştur. CE102, RE102, CS114 ve RS103 testlerini içersin.',
    ic1: 'ICD dokümanına dayalı arayüz test senaryoları üret. Mesaj formatı, zamanlama, hata işleme ve sınır değeri testlerini içersin.',
    ic2: 'Protokol mesaj format doğrulama test adımları oluştur. Nominal, boundary ve negatif test vakalarını içersin.',
    ic3: 'Arayüz zamanlama ve periyot test senaryoları üret. Gecikme, kayıp mesaj, eş zamanlılık ve performans testlerini içersin.',
    ic4: 'Hata enjeksiyon test senaryoları oluştur. Bağlantı kopması, yanlış mesaj, zaman aşımı ve kurtarma testlerini içersin.',
    ic5: 'IRS gereksinim doğrulama test matrisi oluştur. Her arayüz gereksinimi için doğrulama yöntemi ve test case bağlantısını tablola.',
    en1: 'MIL-STD-810H Yöntem 501/502 termal test prosedürü hazırla. Yüksek/düşük sıcaklık, termal şok ve döngü testlerini içersin.',
    en2: 'MIL-STD-810H Yöntem 514 titreşim test planı hazırla. Rastgele, sinüzoidal ve operasyonel titreşim profilleri için test prosedürlerini belirt.',
    en3: 'MIL-STD-810H Yöntem 516 şok test planı oluştur. Yarı-sinüs, terminal peak sawtooth ve trapezoid darbe test prosedürlerini içersin.',
    en4: 'IP67/68 ve MIL-STD-810H Yöntem 506 nem/su geçirmezlik test prosedürü hazırla. Daldırma, yağmur ve nem koşullanma testlerini içersin.',
    en5: 'MIL-STD-810H Yöntem 510 kum ve toz dayanım test prosedürü hazırla. Blown sand, blown dust test konfigürasyonlarını içersin.',
    en6: 'Yazılım çevresel test senaryoları üret. Sıcaklık değişiminde veri bütünlüğü, güç kesintisi kurtarma ve soğuk/sıcak başlatma testlerini içersin.',
    cy1: 'Savunma sistemi siber güvenlik sızma testi planı hazırla. Keşif, zafiyet analizi, istismar ve raporlama aşamalarını içersin.',
    cy2: 'TEMPEST elektromanyetik emanasyon test senaryoları oluştur. Van Eck emanasyonu, compromising emanation ölçüm ve değerlendirmelerini içersin.',
    cy3: 'Kimlik doğrulama ve yetkilendirme test senaryoları üret. Güçlü parola, MFA, oturum yönetimi ve RBAC testlerini içersin.',
    cy4: 'Ağ güvenlik test senaryoları oluştur. Port tarama, servis keşfi, firewall kuralları, şifreleme doğrulama testlerini içersin.',
    cy5: 'Sistem zafiyet analizi ve CVE tarama planı hazırla. Zafiyet önceliklendirme, CVSS skorlama ve düzeltme takip sürecini içersin.',
    fs1: 'Fabrika Kabul Testi (FAT) prosedürü hazırla. Hazırlık, test dizisi, ölçüm, geçiş kriterleri ve raporlama adımlarını içersin.',
    fs2: 'Saha Kabul Testi (SAT) prosedürü hazırla. Saha kurulum, entegrasyon doğrulama, operasyonel test ve kabul adımlarını içersin.',
    fs3: 'FAT/SAT geçiş kriterleri listesi oluştur. Kritik/majör/minör kategorilere göre ölçülebilir kabul şartlarını belirt.',
    fs4: 'FAT hazırlık kontrol listesi oluştur. Doküman kontrolü, ekipman kalibrasyonu, test ortamı ve personel hazırlığını içersin.',
    fs5: 'FAT sonuç raporu formatı oluştur. Yürütülen testler, sonuçlar, anomaliler, açık maddeler ve kabul kararını içersin.',
  }

  const runTools = async() => {
    if(!input.trim()||!selectedTools.length) return
    setLoading(true); setResult('')
    try{
      const toolList = currentTools.filter(t=>selectedTools.includes(t.id))
      let combined = ''
      for(const tool of toolList){
        const prompt = 'Savunma sanayii test mühendisi olarak, asagidaki konuya/sisteme gore "' + tool.label + '" uret:\n\n' + input + '\n\n' + (PROMPTS[tool.id]||'') + '\n\nTurkce, teknik ve uygulanabilir cikti ver. Baslik olarak "' + tool.label + '" kullan.'
        const r = await callAI([{role:'user',content:prompt}])
        combined += (combined ? '\n\n' + '─'.repeat(60) + '\n\n' : '') + r
      }
      setResult(combined)
    }catch(e){setResult('Hata: '+e.message)}
    finally{setLoading(false)}
  }

  const doExport = () => {
    if(!result) return
    const fname = 'savunma_' + activeSection + '_' + new Date().toISOString().slice(0,10)
    if(exportFmt==='txt') downloadFile(result, fname+'.txt', 'text/plain')
    else if(exportFmt==='md') downloadFile('# Savunma Sanayii Ciktisi\n\n' + result, fname+'.md', 'text/markdown')
    else if(exportFmt==='html'){
      const html = '<!DOCTYPE html><html lang="tr"><head><meta charset="UTF-8"><title>Savunma Çıktısı</title><style>body{font-family:Segoe UI,sans-serif;background:#05080f;color:#c8d8e8;padding:32px;max-width:900px;margin:0 auto}h1{color:#f87171}pre{white-space:pre-wrap;line-height:1.8;font-family:Segoe UI,sans-serif;background:rgba(0,0,0,0.3);padding:20px;border-radius:8px;border:1px solid #1a2a4a}</style></head><body><h1>🛡️ Savunma Sanayii Çıktısı</h1><pre>' + result.replace(/</g,'&lt;').replace(/>/g,'&gt;') + '</pre></body></html>'
      downloadFile(html, fname+'.html', 'text/html')
    }
  }

  return(
    <div style={{display:'flex',height:'100%',overflow:'hidden'}}>
      {/* Sol menü */}
      <div style={{width:200,flexShrink:0,borderRight:`1px solid ${C.b0}`,background:C.bg1,overflowY:'auto'}}>
        <div style={{padding:'12px 14px',borderBottom:`1px solid ${C.b0}`}}>
          <div style={{fontSize:13,fontWeight:800,color:C.t0}}>🛡️ Savunma Sanayii</div>
          <div style={{fontSize:10,color:C.t3,marginTop:2}}>Sistem & Yazılım Test Müh.</div>
        </div>
        {sections.map(sec=>(
          <div key={sec.id} onClick={()=>{setActiveSection(sec.id);setSelectedTools([]);setResult('')}}
            style={{padding:'10px 14px',cursor:'pointer',fontSize:12,fontWeight:500,
              color:activeSection===sec.id?sec.color:C.t2,
              background:activeSection===sec.id?`${sec.color}12`:'transparent',
              borderLeft:`3px solid ${activeSection===sec.id?sec.color:'transparent'}`,
              transition:'all .15s'}}
            onMouseEnter={e=>{if(activeSection!==sec.id)e.currentTarget.style.background='rgba(255,255,255,0.02)'}}
            onMouseLeave={e=>{if(activeSection!==sec.id)e.currentTarget.style.background='transparent'}}>
            {sec.label}
          </div>
        ))}
      </div>

      {/* Sağ içerik */}
      <div style={{flex:1,display:'flex',flexDirection:'column',overflow:'hidden'}}>

        {activeSection==='ai_assist' ? (
          <div style={{flex:1,overflowY:'auto',padding:24}}>
            <div style={{fontSize:18,fontWeight:800,color:C.t0,marginBottom:6}}>🤖 Serbest AI Desteği</div>
            <div style={{fontSize:12,color:C.t3,marginBottom:16}}>Herhangi bir savunma mühendisliği sorusu veya belgesi için serbest analiz yap</div>
            <TextArea value={input} onChange={e=>setInput(e.target.value)} rows={6} placeholder="Sistem gereksinimi, ICD özeti, soru veya analiz isteğini buraya yazın..."/>
            <div style={{marginTop:10,display:'flex',gap:8,alignItems:'center'}}>
              <BtnPrimary size="lg" onClick={runTools} loading={loading} disabled={!input.trim()}>🛡️ AI Analizi Yap</BtnPrimary>
            </div>
            {result&&<pre style={{whiteSpace:'pre-wrap',fontSize:13,color:C.t1,lineHeight:1.8,marginTop:16,padding:'16px',background:'rgba(0,0,0,0.2)',borderRadius:9,border:`1px solid ${C.b0}`,fontFamily:"'Segoe UI',system-ui,sans-serif"}}>{result}</pre>}
          </div>
        ) : (
          <>
            {/* Araç seçim paneli */}
            <div style={{padding:'16px 20px',borderBottom:`1px solid ${C.b0}`,background:C.bg1,flexShrink:0}}>
              <div style={{fontSize:12,fontWeight:700,color:C.t2,marginBottom:10}}>
                {sections.find(s=>s.id===activeSection)?.label} — Araçlar
              </div>
              <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fill,minmax(200px,1fr))',gap:7,marginBottom:12}}>
                {currentTools.map(tool=>{
                  const sel = selectedTools.includes(tool.id)
                  const secColor = sections.find(s=>s.id===activeSection)?.color||C.blue
                  return(
                    <div key={tool.id} onClick={()=>toggleTool(tool.id)}
                      style={{display:'flex',alignItems:'center',gap:8,padding:'8px 11px',borderRadius:8,cursor:'pointer',
                        border:`1px solid ${sel?secColor+'60':C.b0}`,
                        background:sel?`${secColor}10`:'rgba(255,255,255,0.01)',
                        transition:'all .15s'}}
                      onMouseEnter={e=>{if(!sel)e.currentTarget.style.background='rgba(255,255,255,0.03)'}}
                      onMouseLeave={e=>{if(!sel)e.currentTarget.style.background='rgba(255,255,255,0.01)'}}>
                      <div style={{width:16,height:16,borderRadius:4,border:`1px solid ${sel?secColor:C.b1}`,background:sel?secColor:'transparent',flexShrink:0,display:'flex',alignItems:'center',justifyContent:'center'}}>
                        {sel&&<span style={{fontSize:10,color:'#fff',lineHeight:1}}>✓</span>}
                      </div>
                      <span style={{fontSize:15,flexShrink:0}}>{tool.icon}</span>
                      <div style={{minWidth:0}}>
                        <div style={{fontSize:12,fontWeight:600,color:sel?sections.find(s=>s.id===activeSection)?.color||C.blue:C.t1,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{tool.label}</div>
                        <div style={{fontSize:10,color:C.t3,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{tool.desc}</div>
                      </div>
                    </div>
                  )
                })}
              </div>
              <TextArea value={input} onChange={e=>setInput(e.target.value)} rows={3}
                placeholder="Sisteminizi, gereksinimi veya analiz etmek istediğiniz konuyu yazın..."/>
              <div style={{marginTop:10,display:'flex',gap:8,alignItems:'center'}}>
                <BtnPrimary onClick={runTools} loading={loading}
                  disabled={!input.trim()||!selectedTools.length||loading}>
                  {loading?<><Spin c={C.blue} s={13}/> Üretiliyor...</>:`▶ ${selectedTools.length||0} Araç Çalıştır`}
                </BtnPrimary>
                {selectedTools.length>0&&<BtnGhost size="sm" onClick={()=>setSelectedTools([])}>Seçimi Temizle</BtnGhost>}
                {result&&(
                  <div style={{display:'flex',gap:6,alignItems:'center',marginTop:4}}>
                    <select value={exportFmt} onChange={e=>setExportFmt(e.target.value)}
                      style={{fontSize:11,padding:'4px 6px',borderRadius:5,border:`1px solid ${C.b1}`,background:C.bg3,color:C.t2,fontFamily:'inherit',cursor:'pointer'}}>
                      <option value="txt">📄 TXT</option>
                      <option value="md">📝 MD</option>
                      <option value="html">🌐 HTML</option>
                    </select>
                    <BtnSuccess size="sm" icon="📥" onClick={doExport}>İndir</BtnSuccess>
                  </div>
                )}
              </div>
            </div>

            {/* Sonuç */}
            <div style={{flex:1,overflowY:'auto',padding:'16px 20px'}}>
              {!result&&!loading&&(
                <div style={{textAlign:'center',padding:'48px 0',color:C.t3}}>
                  <div style={{fontSize:36,marginBottom:10,opacity:.3}}>🛡️</div>
                  <div style={{fontSize:13,fontWeight:700,color:C.t2,marginBottom:5}}>Araç seç → Konuyu yaz → Çalıştır</div>
                  <div style={{fontSize:12}}>Seçilen araçlara göre savunma mühendisliği çıktısı üretilir</div>
                </div>
              )}
              {loading&&!result&&(
                <div style={{textAlign:'center',padding:'48px 0'}}>
                  <Spin c={C.red} s={28}/>
                  <div style={{marginTop:14,fontSize:13,color:C.red}}>Üretiliyor...</div>
                </div>
              )}
              {result&&(
                <pre style={{whiteSpace:'pre-wrap',fontSize:13,color:C.t1,lineHeight:1.85,fontFamily:"'Segoe UI',system-ui,sans-serif",margin:0}}>{result}</pre>
              )}
            </div>
          </>
        )}
      </div>
    </div>
  )
}



// ─── AI ÖĞRENİM TAB ───────────────────────────────────────────────────────────
function LearningTab(){
  const [mode, setMode] = useState('knowledge')
  const [knowledge, setKnowledge] = useState(()=>LS.get('qa_knowledge',[]))
  const [templates, setTemplates] = useState(()=>LS.get('qa_templates',[]))

  // Bilgi modu state
  const [kTitle, setKTitle] = useState('')
  const [kText, setKText] = useState('')
  const [kLoading, setKLoading] = useState(false)
  const [kQuestion, setKQuestion] = useState('')
  const [kAnswer, setKAnswer] = useState('')
  const [kAskLoading, setKAskLoading] = useState(false)
  const kFileRef = useRef()

  // Sablon modu state
  const [tName, setTName] = useState('')
  const [tDoc, setTDoc] = useState('')
  const [tAddLoading, setTAddLoading] = useState(false)
  const [selTmpl, setSelTmpl] = useState(null)
  const [tData, setTData] = useState('')
  const [tResult, setTResult] = useState('')
  const [tRunLoading, setTRunLoading] = useState(false)
  const [tExportFmt, setTExportFmt] = useState('docx')
  const tFileRef = useRef()
  const tDocFileRef = useRef()

  const saveKnowledge = p => { setKnowledge(p); LS.set('qa_knowledge', p) }
  const saveTemplates = p => { setTemplates(p); LS.set('qa_templates', p) }

  const readFile = async file => {
    if(!file) return ''
    try {
      if(file.name.endsWith('.docx')){
        const m = await import('mammoth').catch(()=>null)
        if(m){ const r = await m.extractRawText({arrayBuffer: await file.arrayBuffer()}); return r.value }
      }
      return await file.text()
    } catch(e){ return '' }
  }

  // Bilgi ekle
  const addKnowledge = async () => {
    if(!kText.trim()) return
    setKLoading(true)
    try {
      const summary = await callAI([{role:'user', content:'Asagidaki bilgiyi analiz et, ozetle. Ana kavramlar, kritik noktalar, test ile ilgili cikarimlar. Turkce.\n\n' + kText.slice(0,6000)}])
      const item = {
        id: Date.now(),
        title: kTitle.trim() || kText.slice(0,50) + '...',
        content: kText,
        summary,
        date: new Date().toLocaleDateString('tr-TR')
      }
      saveKnowledge([item, ...knowledge])
      setKTitle(''); setKText('')
    } catch(e){ alert('Hata: ' + e.message) }
    finally{ setKLoading(false) }
  }

  // Bilgi tabanina soru sor
  const askKnowledge = async () => {
    if(!kQuestion.trim() || !knowledge.length) return
    setKAskLoading(true); setKAnswer('')
    try {
      const ctx = knowledge.slice(0,6).map(k => '[' + k.title + ']\n' + k.summary).join('\n\n---\n\n')
      const r = await callAI([{role:'user', content:'Ogrenilmis bilgi tabanina dayanarak yanıtla. Bilgi tabaninda yoksa belirt.\n\nBilgi Tabani:\n' + ctx + '\n\nSoru: ' + kQuestion + '\n\nTurkce, odakli yanıt ver.'}])
      setKAnswer(r)
    } catch(e){ setKAnswer('Hata: ' + e.message) }
    finally{ setKAskLoading(false) }
  }

  // Sablon ekle
  const addTemplate = async () => {
    if(!tDoc.trim() || !tName.trim()) return
    setTAddLoading(true)
    try {
      const analysis = await callAI([{role:'user', content:'Bu dokuman sablonunu analiz et. Yapısını, bolumlerini, doldurulacak alanları ve formatını acikla. Turkce.\n\n' + tDoc.slice(0,6000)}])
      const tmpl = {id: Date.now(), name: tName.trim(), doc: tDoc, analysis, date: new Date().toLocaleDateString('tr-TR')}
      saveTemplates([tmpl, ...templates])
      setTName(''); setTDoc('')
    } catch(e){ alert('Hata: ' + e.message) }
    finally{ setTAddLoading(false) }
  }

  // Sablona veri isle
  const runTemplate = async () => {
    if(!selTmpl || !tData.trim()) return
    setTRunLoading(true); setTResult('')
    try {
      const NL = '\n'
      const prompt = 'GOREV: Sablon dokumandan sadece YAPI ve FORMAT al. Icerik tamamen yeni veriden uretilecek - sablon icerigini kopyalama, sadece yapısini kullan. Tablo varsa tablo yap, baslik varsa baslik, liste varsa liste.' + NL + NL + 'Sablon Yapisi:' + NL + selTmpl.analysis + NL + NL + 'Orijinal Sablon (sadece yapi icin, icerik alma):' + NL + selTmpl.doc.slice(0,4000) + NL + NL + 'Yeni Icerik Verileri:' + NL + tData + NL + NL + 'Sablonun FORMAT/YAPISI kullanilarak tamamen yeni icerikle dolu dokuman uret. Markdown formatinda yaz.'
      const r = await callAI([{role:'user', content: prompt}])
      setTResult(r)
    } catch(e){ setTResult('Hata: ' + e.message) }
    finally{ setTRunLoading(false) }
  }

  const exportTemplate = async () => {
    if(!tResult||!selTmpl) return
    const fname = selTmpl.name.replace(/\s+/g,'_') + '_' + new Date().toISOString().slice(0,10)
    if(tExportFmt==='docx'){
      await exportAsDocx(tResult, selTmpl.name, fname)
    } else if(tExportFmt==='txt'){
      downloadFile(tResult, fname+'.txt', 'text/plain')
    } else if(tExportFmt==='md'){
      downloadFile('# ' + selTmpl.name + '\n\n' + tResult, fname+'.md', 'text/markdown')
    } else if(tExportFmt==='html'){
      const body = tResult.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
      const html = '<!DOCTYPE html><html lang="tr"><head><meta charset="UTF-8"><title>' + selTmpl.name + '</title><style>body{font-family:Segoe UI,sans-serif;line-height:1.8;padding:40px;max-width:860px;margin:0 auto}h1,h2,h3{color:#1e3a8a}table{width:100%;border-collapse:collapse;margin:16px 0}th{background:#dbeafe;padding:8px 12px;text-align:left;border:1px solid #93c5fd}td{padding:8px 12px;border:1px solid #bfdbfe}pre,code{background:#f1f5f9;padding:2px 6px;border-radius:4px;font-family:Consolas,monospace}</style></head><body><h1>' + selTmpl.name + '</h1>' + mdToHtmlSimple(tResult) + '</body></html>'
      downloadFile(html, fname+'.html', 'text/html')
    } else if(tExportFmt==='csv'){
      const rows = tResult.split('\n').map(l=>'"'+l.replace(/"/g,'""')+'"').join('\n')
      downloadFile(rows, fname+'.csv', 'text/csv')
    } else if(tExportFmt==='json'){
      const obj = JSON.stringify({template:selTmpl.name, date:new Date().toISOString(), content:tResult}, null, 2)
      downloadFile(obj, fname+'.json', 'application/json')
    }
  }

  return(
    <div style={{display:'flex',height:'100%',overflow:'hidden'}}>

      {/* Sol panel */}
      <div style={{width:230,flexShrink:0,borderRight:`1px solid ${C.b0}`,background:C.bg1,display:'flex',flexDirection:'column',overflow:'hidden'}}>
        {/* Baslik */}
        <div style={{padding:'14px 16px',borderBottom:`1px solid ${C.b0}`,flexShrink:0}}>
          <div style={{fontSize:13,fontWeight:800,color:C.t0}}>🧠 AI Ogrenim</div>
          <div style={{fontSize:10,color:C.t3,marginTop:2}}>Ogret · Hatirla · Uret</div>
        </div>

        {/* Mod secimi */}
        <div style={{display:'flex',borderBottom:`1px solid ${C.b0}`,flexShrink:0}}>
          {[{id:'knowledge',label:'🧠 Bilgi'},{id:'template',label:'📋 Sablon'}].map(m=>(
            <button key={m.id} onClick={()=>setMode(m.id)}
              style={{flex:1,padding:'9px 6px',border:'none',borderBottom:`2px solid ${mode===m.id?C.purple:'transparent'}`,
                background:'transparent',color:mode===m.id?C.purple:C.t3,cursor:'pointer',fontSize:12,fontWeight:600,fontFamily:'inherit',transition:'all .15s'}}>
              {m.label}
            </button>
          ))}
        </div>

        {/* Liste */}
        <div style={{flex:1,overflowY:'auto',padding:'8px'}}>
          {mode==='knowledge'&&(
            <>
              <div style={{fontSize:9,fontWeight:700,color:C.t3,textTransform:'uppercase',letterSpacing:.8,padding:'4px 6px',marginBottom:4}}>
                Ogrenilen Bilgiler ({knowledge.length})
              </div>
              {knowledge.length===0&&(
                <div style={{textAlign:'center',padding:'24px 8px',color:C.t3,fontSize:11}}>Henuz bilgi eklenmedi</div>
              )}
              {knowledge.map(k=>(
                <div key={k.id} style={{padding:'8px 10px',borderRadius:7,background:'rgba(255,255,255,0.02)',border:`1px solid ${C.b0}`,marginBottom:5}}>
                  <div style={{fontSize:11,fontWeight:600,color:C.t1,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{k.title}</div>
                  <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginTop:4}}>
                    <span style={{fontSize:9,color:C.t3}}>{k.date}</span>
                    <button onClick={()=>saveKnowledge(knowledge.filter(x=>x.id!==k.id))}
                      style={{background:'none',border:'none',color:'#f87171',cursor:'pointer',fontSize:12,padding:'0 2px',lineHeight:1}}>✕</button>
                  </div>
                </div>
              ))}
            </>
          )}
          {mode==='template'&&(
            <>
              <div style={{fontSize:9,fontWeight:700,color:C.t3,textTransform:'uppercase',letterSpacing:.8,padding:'4px 6px',marginBottom:4}}>
                Kaydedilen Sablonlar ({templates.length})
              </div>
              {templates.length===0&&(
                <div style={{textAlign:'center',padding:'24px 8px',color:C.t3,fontSize:11}}>Henuz sablon eklenmedi</div>
              )}
              {templates.map(t=>(
                <div key={t.id} onClick={()=>setSelTmpl(s=>s?.id===t.id?null:t)}
                  style={{padding:'8px 10px',borderRadius:7,marginBottom:5,cursor:'pointer',transition:'all .15s',
                    background:selTmpl?.id===t.id?'rgba(59,158,237,0.1)':'rgba(255,255,255,0.02)',
                    border:`1px solid ${selTmpl?.id===t.id?C.blue:C.b0}`}}>
                  <div style={{fontSize:11,fontWeight:600,color:selTmpl?.id===t.id?C.blue:C.t1,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{t.name}</div>
                  <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginTop:4}}>
                    <span style={{fontSize:9,color:C.t3}}>{t.date}</span>
                    <button onClick={e=>{e.stopPropagation();saveTemplates(templates.filter(x=>x.id!==t.id));if(selTmpl?.id===t.id)setSelTmpl(null)}}
                      style={{background:'none',border:'none',color:'#f87171',cursor:'pointer',fontSize:12,padding:'0 2px',lineHeight:1}}>✕</button>
                  </div>
                </div>
              ))}
            </>
          )}
        </div>
      </div>

      {/* Sag panel */}
      <div style={{flex:1,overflowY:'auto',padding:'24px'}}>

        {/* BİLGİ MODU */}
        {mode==='knowledge'&&(
          <div style={{maxWidth:680}}>
            <div style={{fontSize:17,fontWeight:800,color:C.t0,marginBottom:4}}>🧠 Bilgi Ogrenim</div>
            <div style={{fontSize:12,color:C.t3,marginBottom:22}}>
              Dokuman, prosedur veya teknik bilgi yukle. AI ozetle kaydeder. Daha sonra sorduğunda ogrendiklerini kullanarak yanıt verir.
            </div>

            {/* Bilgi ekleme */}
            <div style={{background:'rgba(139,92,246,0.06)',border:`1px solid rgba(139,92,246,0.25)`,borderRadius:10,padding:'16px 18px',marginBottom:22}}>
              <div style={{fontSize:11,fontWeight:700,color:C.purple,textTransform:'uppercase',letterSpacing:.7,marginBottom:12}}>+ Bilgi / Dokuman Ekle</div>
              <div style={{marginBottom:10}}>
                <Field label="Baslik (opsiyonel)">
                  <TextInput value={kTitle} onChange={e=>setKTitle(e.target.value)} placeholder="Ornek: Sistem Test Proseduru v2.1"/>
                </Field>
              </div>
              <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:6}}>
                <span style={{fontSize:11,fontWeight:600,color:C.t2}}>Icerik</span>
                <button onClick={()=>kFileRef.current.click()}
                  style={{fontSize:10,padding:'3px 9px',borderRadius:5,border:`1px solid ${C.b1}`,background:'rgba(255,255,255,0.03)',color:C.t2,cursor:'pointer',fontFamily:'inherit'}}>
                  📎 Dosya Yukle (TXT / DOCX)
                </button>
                <input ref={kFileRef} type="file" accept=".txt,.md,.docx" style={{display:'none'}}
                  onChange={async e=>{ const t=await readFile(e.target.files[0]); if(t) setKText(p=>p?p+'\n\n'+t:t); e.target.value='' }}/>
              </div>
              <TextArea value={kText} onChange={e=>setKText(e.target.value)} rows={5}
                placeholder="Ogretmek istedigin bilgiyi, dokumani veya konuyu buraya yapistir..."/>
              <div style={{marginTop:12}}>
                <BtnAccent onClick={addKnowledge} loading={kLoading} disabled={!kText.trim()} icon="🧠">
                  {kLoading ? 'Ogreniyor...' : "AI'ya Ogret"}
                </BtnAccent>
              </div>
            </div>

            {/* Soru sorma */}
            {knowledge.length>0&&(
              <div style={{background:'rgba(59,158,237,0.06)',border:`1px solid rgba(59,158,237,0.25)`,borderRadius:10,padding:'16px 18px'}}>
                <div style={{fontSize:11,fontWeight:700,color:C.blue,textTransform:'uppercase',letterSpacing:.7,marginBottom:8}}>
                  💬 Ogrenilenlere Dayanarak Soru Sor
                </div>
                <div style={{fontSize:11,color:C.t3,marginBottom:12}}>
                  AI {knowledge.length} ogrenilmis bilgiyi baz alarak yanit verecek
                </div>
                <div style={{display:'flex',gap:8,marginBottom:12}}>
                  <TextInput value={kQuestion} onChange={e=>setKQuestion(e.target.value)}
                    onKeyDown={e=>e.key==='Enter'&&askKnowledge()}
                    placeholder="Ogrenilen bilgiler hakkinda soru sor..."/>
                  <BtnPrimary onClick={askKnowledge} loading={kAskLoading} disabled={!kQuestion.trim()||kAskLoading}>Sor</BtnPrimary>
                </div>
                {kAnswer&&(
                  <div style={{padding:'14px',background:'rgba(0,0,0,0.2)',borderRadius:8,border:`1px solid ${C.b0}`}}>
                    <pre style={{whiteSpace:'pre-wrap',fontSize:13,color:C.t1,lineHeight:1.8,fontFamily:"'Segoe UI',system-ui,sans-serif",margin:0}}>{kAnswer}</pre>
                  </div>
                )}
              </div>
            )}
            {knowledge.length===0&&(
              <div style={{textAlign:'center',padding:'30px 0',color:C.t3}}>
                <div style={{fontSize:32,marginBottom:10,opacity:.3}}>🧠</div>
                <div style={{fontSize:13,color:C.t2,fontWeight:600,marginBottom:5}}>Henuz bilgi eklenmedi</div>
                <div style={{fontSize:12}}>Yukardaki formdan bilgi veya dokuman ekle</div>
              </div>
            )}
          </div>
        )}

        {/* SABLON MODU */}
        {mode==='template'&&(
          <div style={{maxWidth:680}}>
            <div style={{fontSize:17,fontWeight:800,color:C.t0,marginBottom:4}}>📋 Sablon Ogrenim</div>
            <div style={{fontSize:12,color:C.t3,marginBottom:22}}>
              Bir dokuman sablonu yukle, AI yapısını ogrenir. Sonra veri girince o sablona uygun cikti uretir.
            </div>

            {/* Sablon ekleme */}
            {!selTmpl&&(
              <div style={{background:'rgba(16,185,129,0.06)',border:`1px solid rgba(16,185,129,0.25)`,borderRadius:10,padding:'16px 18px',marginBottom:20}}>
                <div style={{fontSize:11,fontWeight:700,color:C.green,textTransform:'uppercase',letterSpacing:.7,marginBottom:12}}>+ Yeni Sablon Ogret</div>
                <div style={{marginBottom:10}}>
                  <Field label="Sablon Adi">
                    <TextInput value={tName} onChange={e=>setTName(e.target.value)} placeholder="Ornek: Test Raporu Sablonu v3"/>
                  </Field>
                </div>
                <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:6}}>
                  <span style={{fontSize:11,fontWeight:600,color:C.t2}}>Sablon Dokumani</span>
                  <button onClick={()=>tDocFileRef.current.click()}
                    style={{fontSize:10,padding:'3px 9px',borderRadius:5,border:`1px solid ${C.b1}`,background:'rgba(255,255,255,0.03)',color:C.t2,cursor:'pointer',fontFamily:'inherit'}}>
                    📎 Dosya Yukle (TXT / DOCX)
                  </button>
                  <input ref={tDocFileRef} type="file" accept=".txt,.md,.docx" style={{display:'none'}}
                    onChange={async e=>{ const t=await readFile(e.target.files[0]); if(t) setTDoc(t); e.target.value='' }}/>
                </div>
                <TextArea value={tDoc} onChange={e=>setTDoc(e.target.value)} rows={6}
                  placeholder="Sablon dokumani buraya yapistir veya yukle. AI bu sablonun yapısını ogrenecek..."/>
                <div style={{marginTop:12}}>
                  <BtnSuccess onClick={addTemplate} loading={tAddLoading} disabled={!tDoc.trim()||!tName.trim()} icon="📋">
                    {tAddLoading ? 'Analiz ediliyor...' : 'Sablonu Ogret'}
                  </BtnSuccess>
                </div>
              </div>
            )}

            {/* Secili sablon ile veri isleme */}
            {selTmpl&&(
              <div style={{background:'rgba(59,158,237,0.06)',border:`1px solid rgba(59,158,237,0.25)`,borderRadius:10,padding:'16px 18px',marginBottom:20}}>
                <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:14}}>
                  <div style={{fontSize:13,fontWeight:700,color:C.blue,flex:1}}>📋 {selTmpl.name}</div>
                  <BtnGhost size="sm" onClick={()=>{setSelTmpl(null);setTResult('')}}>✕ Kapat</BtnGhost>
                </div>

                {/* Sablon analizi ozeti */}
                <div style={{padding:'10px 12px',borderRadius:7,background:'rgba(0,0,0,0.2)',border:`1px solid ${C.b0}`,marginBottom:14,maxHeight:110,overflowY:'auto'}}>
                  <div style={{fontSize:9,color:C.t3,fontWeight:700,textTransform:'uppercase',letterSpacing:.7,marginBottom:4}}>AI Sablon Analizi</div>
                  <pre style={{whiteSpace:'pre-wrap',fontSize:11,color:C.t2,lineHeight:1.6,fontFamily:"'Segoe UI',system-ui,sans-serif",margin:0}}>{selTmpl.analysis.slice(0,500)}{selTmpl.analysis.length>500?'...':''}</pre>
                </div>

                <div style={{marginBottom:12}}>
                  <Field label="Islenecek Veri / Bilgiler">
                    <TextArea value={tData} onChange={e=>setTData(e.target.value)} rows={5}
                      placeholder="Sablona islenecek verileri gir. Ornek: Sistem adi, test tarihi, sonuclar, bulgular, sorumlu kisi..."/>
                  </Field>
                </div>

                <div style={{display:'flex',gap:8,alignItems:'center',flexWrap:'wrap'}}>
                  <BtnPrimary onClick={runTemplate} loading={tRunLoading} disabled={!tData.trim()||tRunLoading}>
                    {tRunLoading ? <><Spin c={C.blue} s={13}/> Uretiliyor...</> : '📋 Sablona Isle'}
                  </BtnPrimary>
                  {tResult&&(
                    <div style={{display:'flex',gap:6,alignItems:'center',flexWrap:'wrap',marginTop:4}}>
                      <select value={tExportFmt} onChange={e=>setTExportFmt(e.target.value)}
                        style={{fontSize:11,padding:'5px 8px',borderRadius:6,border:`1px solid ${C.b1}`,background:C.bg2,color:C.t1,fontFamily:'inherit',cursor:'pointer'}}>
                        <option value="docx">📘 Word (.docx)</option>
                        <option value="txt">📄 Text (.txt)</option>
                        <option value="md">📝 Markdown (.md)</option>
                        <option value="html">🌐 HTML (.html)</option>
                        <option value="csv">📊 CSV (.csv)</option>
                        <option value="json">🔧 JSON (.json)</option>
                      </select>
                      <BtnSuccess size="sm" icon="📥" onClick={exportTemplate}>Indir</BtnSuccess>
                      <BtnGhost size="sm" onClick={()=>navigator.clipboard.writeText(tResult)}>Kopyala</BtnGhost>
                    </div>
                  )}
                </div>

                {tResult&&(
                  <div style={{marginTop:16,padding:'16px',background:'rgba(0,0,0,0.2)',borderRadius:8,border:`1px solid ${C.b0}`}}>
                    <pre style={{whiteSpace:'pre-wrap',fontSize:13,color:C.t1,lineHeight:1.8,fontFamily:"'Segoe UI',system-ui,sans-serif",margin:0}}>{tResult}</pre>
                  </div>
                )}
              </div>
            )}

            {!selTmpl&&templates.length>0&&(
              <div style={{textAlign:'center',padding:'20px 0',color:C.t3,fontSize:12}}>
                ← Sol panelden bir sablon sec
              </div>
            )}
            {!selTmpl&&templates.length===0&&(
              <div style={{textAlign:'center',padding:'30px 0',color:C.t3}}>
                <div style={{fontSize:32,marginBottom:10,opacity:.3}}>📋</div>
                <div style={{fontSize:13,color:C.t2,fontWeight:600,marginBottom:5}}>Henuz sablon eklenmedi</div>
                <div style={{fontSize:12}}>Yukardaki formdan sablon yukle</div>
              </div>
            )}
          </div>
        )}

      </div>
    </div>
  )
}


// ─── ASSISTANT ────────────────────────────────────────────────────────────────
function AssistantTab(){
  const [sessions,setSessions]=useState(()=>LS.get('qa_chat_sessions',[{id:Date.now(),title:'Yeni Sohbet',messages:[{role:'assistant',content:'Merhaba! 👋 QA Studio AI Asistanıyım.\n\nTest mühendisliği, otomasyon stratejileri, hata analizi veya araç tavsiyeleri konusunda yardımcı olabilirim.'}]}]))
  const [activeId,sActive]=useState(()=>LS.get('qa_chat_sessions',[])[0]?.id||Date.now())
  const [input,sInput]=useState('')
  const [loading,sLoading]=useState(false)
  const bottom=useRef()
  const active=sessions.find(s=>s.id===activeId)||sessions[0]
  const msgs=active?.messages||[]
  useEffect(()=>{ bottom.current?.scrollIntoView({behavior:'smooth'}) },[msgs,loading])
  useEffect(()=>{ LS.set('qa_chat_sessions',sessions) },[sessions])
  const updMsgs=(id,m)=>setSessions(ss=>ss.map(s=>s.id===id?{...s,messages:m,title:m.find(x=>x.role==='user')?.content?.slice(0,30)||s.title}:s))
  const newSession=()=>{ const s={id:Date.now(),title:'Yeni Sohbet',messages:[{role:'assistant',content:'Yeni sohbet başladı. Ne öğrenmek ya da sormak istersin?'}]}; setSessions(ss=>[s,...ss]); sActive(s.id) }
  const delSession=id=>{ const u=sessions.filter(s=>s.id!==id); setSessions(u.length?u:[{id:Date.now(),title:'Sohbet',messages:[{role:'assistant',content:'Merhaba!'}]}]); if(activeId===id)sActive(u[0]?.id) }
  const send=async()=>{
    if(!input.trim()||loading) return
    const um={role:'user',content:input}
    const nm=[...msgs,um]; updMsgs(activeId,nm); sInput(''); sLoading(true)
    try{ const r=await callAI(nm.filter(m=>m.role==='user'||m.role==='assistant').map(m=>({role:m.role,content:m.content}))); updMsgs(activeId,[...nm,{role:'assistant',content:r}]) }
    catch(e){ updMsgs(activeId,[...nm,{role:'assistant',content:'⚠️ Hata: '+e.message}]) }
    finally{ sLoading(false) }
  }
  const hints=['Playwright nasıl kurulur?','OWASP Top 10 nedir?','Coverage nasıl artırılır?','İyi bug raporu nasıl yazılır?','k6 yük testi nasıl yapılır?','API mock ne zaman kullanılır?']
  return(
    <div style={{display:'flex',height:'100%',background:C.bg0}}>
      {/* Sidebar */}
      <div style={{width:210,borderRight:`1px solid ${C.b0}`,background:C.bg1,display:'flex',flexDirection:'column',flexShrink:0}}>
        <div style={{padding:'12px 10px'}}>
          <BtnPrimary onClick={newSession} sx={{width:'100%',justifyContent:'center'}}>+ Yeni Sohbet</BtnPrimary>
        </div>
        <div style={{flex:1,overflowY:'auto',padding:'0 8px 12px'}}>
          {sessions.map(s=>(
            <div key={s.id} onClick={()=>sActive(s.id)} style={{display:'flex',alignItems:'center',gap:6,padding:'8px 10px',borderRadius:7,marginBottom:2,cursor:'pointer',background:s.id===activeId?'rgba(59,158,237,0.1)':'transparent',border:`1px solid ${s.id===activeId?C.b2:'transparent'}`,transition:'all .15s'}}>
              <span style={{flex:1,fontSize:12,color:s.id===activeId?C.blue:C.t3,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>💬 {s.title}</span>
              <button onClick={e=>{e.stopPropagation();delSession(s.id)}} style={{background:'none',border:'none',color:C.t3,cursor:'pointer',fontSize:12,flexShrink:0,padding:0}} onMouseEnter={e=>e.currentTarget.style.color=C.t1} onMouseLeave={e=>e.currentTarget.style.color=C.t3}>✕</button>
            </div>
          ))}
        </div>
      </div>
      {/* Chat */}
      <div style={{flex:1,display:'flex',flexDirection:'column',overflow:'hidden'}}>
        <div style={{flex:1,overflowY:'auto',padding:'20px 24px',display:'flex',flexDirection:'column',gap:12}}>
          {msgs.map((m,i)=>(
            <div key={i} style={{display:'flex',justifyContent:m.role==='user'?'flex-end':'flex-start'}}>
              <div style={{maxWidth:'72%',padding:'11px 15px',borderRadius:m.role==='user'?'12px 12px 3px 12px':'12px 12px 12px 3px',background:m.role==='user'?'rgba(59,158,237,0.18)':C.bg2,border:`1px solid ${m.role==='user'?C.b2:C.b0}`,color:C.t1,fontSize:13,lineHeight:1.75}}>
                {m.role==='assistant'&&<div style={{fontSize:10,color:C.blue,fontWeight:700,marginBottom:6,textTransform:'uppercase',letterSpacing:.7}}>🤖 QA Asistan</div>}
                <pre style={{whiteSpace:'pre-wrap',wordBreak:'break-word',margin:0,fontFamily:"'Segoe UI',system-ui,sans-serif",fontSize:13}}>{m.content}</pre>
              </div>
            </div>
          ))}
          {loading&&<div style={{display:'flex'}}><div style={{padding:'11px 15px',borderRadius:'12px 12px 12px 3px',background:C.bg2,border:`1px solid ${C.b0}`,display:'flex',alignItems:'center',gap:9}}><Spin/><span style={{color:C.t3,fontSize:13}}>Yanıt yazılıyor...</span></div></div>}
          {msgs.length<=1&&(
            <div style={{display:'flex',gap:6,flexWrap:'wrap'}}>
              {hints.map(h=>(
                <button key={h} onClick={()=>sInput(h)} style={{padding:'6px 12px',borderRadius:20,border:`1px solid ${C.b1}`,background:'transparent',color:C.t3,cursor:'pointer',fontSize:12,fontFamily:'inherit',transition:'all .15s'}}
                  onMouseEnter={e=>{e.currentTarget.style.borderColor=C.blue;e.currentTarget.style.color=C.blue}}
                  onMouseLeave={e=>{e.currentTarget.style.borderColor=C.b1;e.currentTarget.style.color=C.t3}}>
                  {h}
                </button>
              ))}
            </div>
          )}
          <div ref={bottom}/>
        </div>
        <div style={{padding:'12px 20px',borderTop:`1px solid ${C.b0}`,background:C.bg1,display:'flex',gap:10,alignItems:'flex-end'}}>
          <TextArea value={input} onChange={e=>sInput(e.target.value)} onKeyDown={e=>{if(e.key==='Enter'&&!e.shiftKey){e.preventDefault();send()}}} placeholder="Bir şey sor... (Enter ile gönder, Shift+Enter yeni satır)" rows={3} sx={{flex:1}}/>
          <BtnPrimary onClick={send} disabled={!input.trim()||loading} size="lg" sx={{padding:'11px 16px',alignSelf:'flex-end'}}>{loading?<Spin c="#fff"/>:'↑'}</BtnPrimary>
        </div>
      </div>
    </div>
  )
}

// ─── MAIN APP ─────────────────────────────────────────────────────────────────
export default function Home(){
  const [splash,setSplash]=useState(true)
  const [splashOut,setSplashOut]=useState(false)
  const [tab,setTab]=useState('studio')
  const [tabTransition,setTabTransition]=useState(false)
  const [pendingTab,setPendingTab]=useState(null)
  const [customTools,setCT]=useState([])
  const [presets,setPresets]=useState(()=>LS.get('qa_presets', null))
  const [confirm,setConfirm]=useState(null) // {message, onConfirm}
  const [presetEdit,setPresetEdit]=useState(false)
  const [bugs,setBugs]=useState([])
  const [selectedTools,setST]=useState([])
  const [input,setInput]=useState('')
  const [results,setResults]=useState([])  // her eleman: {tool, content, messages:[{role,content}]}
  const [running,setRunning]=useState(false)
  const [progress,setProgress]=useState(null)
  const [catFilter,setCF]=useState('Tümü')
  const [search,setSearch]=useState('')
  const [error,setError]=useState(null)
  const [hiddenTools,setHiddenTools]=useState(()=>LS.get('qa_hidden_tools',[]))
  const [editDefaultTool,setEditDefaultTool]=useState(null) // varsayılan araç düzenleme
  const [toolOrder,setToolOrder]=useState(()=>LS.get('qa_tool_order',null))
  const [usageStats,setUS]=useState({})
  const [icdContent,setIcd]=useState('')

  useEffect(()=>{
    try {
      const saved = localStorage.getItem('qa_custom_tools')
      if(saved) setCT(JSON.parse(saved))
    } catch(e) { setCT([]) }
    setBugs(LS.get('qa_bugs',[]))
    setUS(LS.get('qa_usage_stats',{}))
    setIcd(LS.get('qa_icd_content',''))
  },[])

  const editedTools = LS.get('qa_edited_tools',{})
  const allTools=[...DEFAULT_TOOLS.map(t=>editedTools[t.id]?{...t,...editedTools[t.id]}:t),...customTools]
  const askConfirm = (message, onConfirm) => setConfirm({message, onConfirm})
  const allPresets = presets || PRESETS
  const savePresets = (p) => { setPresets(p); LS.set('qa_presets', p) }
  const topIds=Object.entries(usageStats).sort((a,b)=>b[1]-a[1]).slice(0,5).map(([id])=>id)
  const toggle=id=>setST(s=>s.includes(id)?s.filter(x=>x!==id):[...s,id])

  const trackUsage=ids=>setUS(prev=>{
    const u={...prev}; ids.forEach(id=>{u[id]=(u[id]||0)+1})
    LS.set('qa_usage_stats',u); return u
  })

  const orderedTools = toolOrder
    ? [...allTools].sort((a,b)=>{ const oi=toolOrder.indexOf(a.id), oj=toolOrder.indexOf(b.id); return (oi===-1?999:oi)-(oj===-1?999:oj) })
    : allTools
  const filtered=orderedTools.filter(t=>!hiddenTools.includes(t.id)).filter(t=>{
    const ok=catFilter==='Tümü'||t.cat===catFilter
    const q=search.toLowerCase()
    return ok&&(!search||t.label.toLowerCase().includes(q)||t.desc.toLowerCase().includes(q))
  })

  const runAll=async()=>{
    if(!selectedTools.length||!input.trim()||running) return
    setRunning(true); setResults([]); setError(null)
    const tools=allTools.filter(t=>selectedTools.includes(t.id))
    trackUsage(selectedTools)
    const icdPfx=icdContent?`ICD bilgilerini dikkate al:\n\n${icdContent.slice(0,3000)}\n\n---\n\n`:''
    try{
      for(let i=0;i<tools.length;i++){
        setProgress(`${i+1}/${tools.length}`)
        const isCodeTool = tools[i].toolType && CODE_TYPES.includes(tools[i].toolType)
        const codePfx = isCodeTool ? `Sadece ${TOOL_TYPES[tools[i].toolType]?.label} kodu üret. Açıklama ekleme, markdown kod bloğu içinde döndür.\n\n` : ''
        const content=await callAI([{role:'user',content:codePfx+icdPfx+buildPrompt(tools[i].id,input,tools[i].prompt)}])
        setResults(r=>[...r,{tool:tools[i],content,messages:[{role:'user',content:input},{role:'assistant',content}]}])
      }
    }catch(e){setError(e.message)}
    finally{setRunning(false);setProgress(null)}
  }

  const canRun=selectedTools.length>0&&input.trim()&&!running

  const enterApp=()=>{
    setSplashOut(true)
    // Animasyon süresiyle tam eşleşsin, takılma olmasın
    setTimeout(()=>setSplash(false), 1350)
  }

  const switchTab=(id)=>{
    if(id===tab) return
    setTabTransition(true)
    setTimeout(()=>{
      setTab(id)
      setTabTransition(false)
    }, 180)
  }

  const TABS=[
    {id:'studio',   label:'🛠',  text:'Test Araçları'},
    {id:'dashboard',label:'📊',  text:'Dashboard'},
    {id:'manager',  label:'⚙️',  text:'Araç Yöneticisi'},
    {id:'assistant',label:'🤖',  text:'AI Asistan'},
    {id:'process',  label:'🔀',  text:'Test Süreçleri'},
    {id:'defense',  label:'🛡️', text:'Savunma Sanayii'},
    {id:'learning', label:'🧠',  text:'AI Öğrenim'},
  ]
  const [navCollapsed,setNavCollapsed]=useState(false)
  const [mobileSidebar,setMobileSidebar]=useState(false)

  // Splash aktifken app arkada hazır bekliyor (görünmez), geçiş kesintisiz
  const appVisible = !splash

  return(
    <>
      {/* App — splash geçişi biterken arkada hazır */}
      <div style={{height:'100vh',background:C.bg0,opacity:appVisible?1:0,transition:appVisible?'opacity .3s ease .1s':'none',pointerEvents:appVisible?'auto':'none',color:C.t1,fontFamily:"'Segoe UI',system-ui,-apple-system,sans-serif",display:'flex',flexDirection:'column',overflow:'hidden'}}>
      <style>{`
        @keyframes qa-spin{to{transform:rotate(360deg)}}
        @media (max-width:768px){
          .qa-sidebar{width:0!important;overflow:hidden!important;border:none!important}
          .qa-mobile-menu-btn{display:flex!important}
          .qa-mobile-overlay{display:block!important}
        }
        @media (max-width:480px){
          .qa-header-nav{display:none!important}
        }
        @keyframes tabOut { 0%{opacity:1;transform:scale(1);filter:blur(0px)} 100%{opacity:0;transform:scale(1.12);filter:blur(14px)} }
        @keyframes tabIn  { 0%{opacity:0} 100%{opacity:1} }
        *{box-sizing:border-box;margin:0;padding:0}
        ::-webkit-scrollbar{width:5px;height:5px}
        ::-webkit-scrollbar-track{background:transparent}
        ::-webkit-scrollbar-thumb{background:rgba(255,255,255,0.08);border-radius:3px}
        ::-webkit-scrollbar-thumb:hover{background:rgba(255,255,255,0.14)}
        ::placeholder{color:rgba(120,143,168,0.5)}
      `}</style>

      {/* ── TOPBAR ── */}
      <header style={{height:52,background:C.bg1,borderBottom:`1px solid ${C.b0}`,display:'flex',alignItems:'center',padding:'0 16px',gap:8,flexShrink:0,zIndex:10}}>
        {/* Mobil menü butonu */}
        <button onClick={()=>setMobileSidebar(v=>!v)}
          style={{display:'none',width:34,height:34,borderRadius:7,border:`1px solid ${C.b0}`,background:'transparent',color:C.t2,cursor:'pointer',fontSize:16,alignItems:'center',justifyContent:'center',flexShrink:0}}
          className="qa-mobile-menu-btn">☰</button>
        {/* Logo */}
        <div style={{display:'flex',alignItems:'center',gap:7,paddingRight:12,borderRight:`1px solid ${C.b0}`}}>
          <span style={{fontWeight:900,fontSize:15,letterSpacing:-.5,background:'linear-gradient(135deg,#fbbf24,#f59e0b)',WebkitBackgroundClip:'text',WebkitTextFillColor:'transparent'}}>QAI</span>
          <span style={{fontWeight:800,fontSize:15,color:C.t0,letterSpacing:-.3,marginLeft:4}}>Studio</span>
        </div>

        {/* Nav tabs — scrollable + collapse */}
        <div style={{display:'flex',alignItems:'center',gap:4,flex:1,minWidth:0}}>
          {/* Collapse butonu */}
          <button onClick={()=>setNavCollapsed(v=>!v)} title={navCollapsed?'Menüyü Aç':'Menüyü Kapat'}
            style={{flexShrink:0,width:26,height:26,borderRadius:6,border:`1px solid ${C.b0}`,background:'transparent',color:C.t3,cursor:'pointer',fontSize:11,display:'flex',alignItems:'center',justifyContent:'center',transition:'all .15s'}}
            onMouseEnter={e=>{e.currentTarget.style.color=C.t1;e.currentTarget.style.background=C.bg3}}
            onMouseLeave={e=>{e.currentTarget.style.color=C.t3;e.currentTarget.style.background='transparent'}}>
            {navCollapsed?'▶':'◀'}
          </button>

          {/* Açık mod — sol/sağ butonlu kaydırma */}
          {!navCollapsed&&(
            <div style={{display:'flex',alignItems:'center',flex:1,minWidth:0,position:'relative'}}>
              {/* Sol ok */}
              <button
                onClick={()=>{const el=document.getElementById('qa-nav');if(el)el.scrollBy({left:-120,behavior:'smooth'})}}
                style={{flexShrink:0,width:22,height:28,borderRadius:5,border:`1px solid ${C.b0}`,background:C.bg2,color:C.t3,cursor:'pointer',fontSize:11,display:'flex',alignItems:'center',justifyContent:'center',marginRight:2,zIndex:1}}
                onMouseEnter={e=>{e.currentTarget.style.color=C.t1;e.currentTarget.style.background=C.bg3}}
                onMouseLeave={e=>{e.currentTarget.style.color=C.t3;e.currentTarget.style.background=C.bg2}}>‹</button>
              <nav id="qa-nav"
                style={{display:'flex',gap:2,overflowX:'auto',flex:1,minWidth:0,scrollbarWidth:'none',msOverflowStyle:'none',WebkitOverflowScrolling:'touch'}}
                onWheel={e=>{e.preventDefault();e.currentTarget.scrollBy({left:e.deltaY>0?80:-80,behavior:'smooth'})}}>
                <style>{`#qa-nav::-webkit-scrollbar{display:none}`}</style>
                {TABS.map(t=>(
                  <button key={t.id} onClick={()=>switchTab(t.id)}
                    style={{display:'flex',alignItems:'center',gap:5,padding:'5px 11px',borderRadius:7,
                      border:`1px solid ${tab===t.id?C.b2:'transparent'}`,
                      background:tab===t.id?'rgba(59,158,237,0.1)':'transparent',
                      color:tab===t.id?C.blue:C.t3,
                      cursor:'pointer',fontSize:12,fontWeight:600,fontFamily:'inherit',
                      transition:'all .15s',whiteSpace:'nowrap',flexShrink:0}}
                    onMouseEnter={e=>{if(tab!==t.id){e.currentTarget.style.color=C.t2;e.currentTarget.style.background='rgba(255,255,255,0.03)'}}}
                    onMouseLeave={e=>{if(tab!==t.id){e.currentTarget.style.color=C.t3;e.currentTarget.style.background='transparent'}}}>
                    <span style={{fontSize:14}}>{t.label}</span>
                    <span className="qa-nav-text">{t.text}</span>
                    {t.id==='manager'&&customTools.length>0&&<Tag color={C.blue}>{customTools.length}</Tag>}
                  </button>
                ))}
              </nav>
              {/* Sağ ok */}
              <button
                onClick={()=>{const el=document.getElementById('qa-nav');if(el)el.scrollBy({left:120,behavior:'smooth'})}}
                style={{flexShrink:0,width:22,height:28,borderRadius:5,border:`1px solid ${C.b0}`,background:C.bg2,color:C.t3,cursor:'pointer',fontSize:11,display:'flex',alignItems:'center',justifyContent:'center',marginLeft:2}}
                onMouseEnter={e=>{e.currentTarget.style.color=C.t1;e.currentTarget.style.background=C.bg3}}
                onMouseLeave={e=>{e.currentTarget.style.color=C.t3;e.currentTarget.style.background=C.bg2}}>›</button>
            </div>
          )}

          {/* Kapalı mod — sadece ikonlar */}
          {navCollapsed&&(
            <div style={{display:'flex',gap:2,flexShrink:0}}>
              {TABS.map(t=>(
                <button key={t.id} onClick={()=>switchTab(t.id)} title={t.text}
                  style={{width:32,height:32,borderRadius:7,
                    border:`1px solid ${tab===t.id?C.b2:'transparent'}`,
                    background:tab===t.id?'rgba(59,158,237,0.1)':'transparent',
                    color:tab===t.id?C.blue:C.t3,
                    cursor:'pointer',fontSize:15,fontFamily:'inherit',
                    transition:'all .15s',display:'flex',alignItems:'center',justifyContent:'center'}}>
                  {t.label}
                </button>
              ))}
            </div>
          )}
        </div>

        {/* Context actions */}
        {tab==='studio'&&(results.length>0||selectedTools.length>0)&&(
          <div style={{display:'flex',alignItems:'center',gap:6,paddingLeft:8,borderLeft:`1px solid ${C.b0}`}}>
            {results.length>0&&<BtnSuccess size="sm" icon="📥" onClick={()=>exportTXT(results)}>TXT İndir</BtnSuccess>}
            <BtnGhost size="sm" onClick={()=>{setST([]);setResults([]);setInput('');setError(null)}}>✕ Sıfırla</BtnGhost>
          </div>
        )}
      </header>

      {/* ── STUDIO ── */}
      {tab==='studio'&&(
        <div style={{display:'flex',flex:1,overflow:'hidden',animation:tabTransition?'tabOut .18s ease forwards':'tabIn .18s ease'}}>

          {/* Sidebar */}
          {/* Mobil overlay */}
          {mobileSidebar&&<div onClick={()=>setMobileSidebar(false)} style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.5)',zIndex:99,display:'none'}} className="qa-mobile-overlay"/>}
          <aside className="qa-sidebar" style={{width:268,flexShrink:0,borderRight:`1px solid ${C.b0}`,background:C.bg1,display:'flex',flexDirection:'column',overflow:'hidden'}}>
            {/* Search */}
            <div style={{padding:'10px 10px 8px'}}>
              <TextInput value={search} onChange={e=>setSearch(e.target.value)} placeholder="Araç ara..." sx={{fontSize:12}}/>
            </div>

            {/* Sık kullanılan */}
            {topIds.length>0&&<SikKullanilanPanel topIds={topIds} allTools={allTools} selectedTools={selectedTools} toggle={toggle}/>}

            {/* Kategori filtreler */}
            <div style={{padding:'0 10px 8px',display:'flex',flexWrap:'wrap',gap:3}}>
              {['Tümü',...CATEGORIES].map(c=>(
                <button key={c} onClick={()=>setCF(c)} style={{padding:'3px 8px',borderRadius:5,fontSize:11,fontWeight:600,border:`1px solid ${catFilter===c?C.b2:C.b0}`,background:catFilter===c?'rgba(59,158,237,0.08)':'transparent',color:catFilter===c?C.blue:C.t3,cursor:'pointer',fontFamily:'inherit',transition:'all .15s'}}>{c}</button>
              ))}
            </div>

            <Divider/>

            {/* Tool list */}
            <div style={{flex:1,overflowY:'auto',padding:'4px 6px 12px'}}>
              <AccordionToolList categories={CATEGORIES} catFilter={catFilter} filtered={filtered} selectedTools={selectedTools} toggle={toggle}
                onMoveUp={(id)=>{
                  const ids=orderedTools.map(t=>t.id)
                  const i=ids.indexOf(id); if(i<=0) return
                  [ids[i-1],ids[i]]=[ids[i],ids[i-1]]
                  setToolOrder(ids); LS.set('qa_tool_order',ids)
                }}
                onMoveDown={(id)=>{
                  const ids=orderedTools.map(t=>t.id)
                  const i=ids.indexOf(id); if(i>=ids.length-1) return
                  [ids[i],ids[i+1]]=[ids[i+1],ids[i]]
                  setToolOrder(ids); LS.set('qa_tool_order',ids)
                }}
                onEditDefault={(tool)=>setEditDefaultTool({...tool})}
                onDeleteTool={(tool)=>askConfirm(`"${tool.label}" aracını kaldırmak istediğinize emin misiniz?`, ()=>{ const isCustom=customTools.find(x=>x.id===tool.id); if(isCustom){const u=customTools.filter(x=>x.id!==tool.id);setCT(u);LS.set('qa_custom_tools',u)} else{const hidden=LS.get('qa_hidden_tools',[]); LS.set('qa_hidden_tools',[...hidden,tool.id]); setHiddenTools(h=>[...h,tool.id])} setConfirm(null)})}/>
            </div>
          </aside>

          {/* Main content */}
          <main style={{flex:1,display:'flex',flexDirection:'column',overflow:'hidden'}}>
            {/* Input panel */}
            <div style={{padding:'14px 18px',borderBottom:`1px solid ${C.b0}`,background:C.bg1,flexShrink:0}}>

              <ICDPanel icdContent={icdContent} setIcdContent={setIcd}/>

              {/* Presets */}
              <PresetPanel presets={allPresets} allTools={allTools} selectedTools={selectedTools} setST={setST} savePresets={savePresets}/>

              <TextArea value={input} onChange={e=>setInput(e.target.value)} placeholder={"Konunu yaz...\n\nÖrnek: Kullanıcıların giriş yapıp profil düzenleyebildiği React web uygulaması"} rows={4}/>

              {/* Seçili araçlar */}
              {selectedTools.length>0&&(
                <div style={{marginTop:10,display:'flex',flexWrap:'wrap',gap:5,padding:'8px 10px',background:'rgba(255,255,255,0.02)',borderRadius:8,border:`1px solid ${C.b0}`,alignItems:'center'}}>
                  <span style={{fontSize:10,fontWeight:700,color:C.t3,textTransform:'uppercase',letterSpacing:.8,flexShrink:0}}>Seçili:</span>
                  {selectedTools.map(id=>{
                    const t=allTools.find(x=>x.id===id); if(!t) return null
                    const cl=PALETTE[t.color]||PALETTE.blue
                    return <Pill key={id} color={cl.text} onClick={()=>toggle(id)} active>{t.icon} {t.label} ✕</Pill>
                  })}
                  <span style={{marginLeft:'auto',cursor:'pointer',fontSize:11,color:C.t3}} onClick={()=>setST([])}>Tümünü kaldır</span>
                </div>
              )}
              <div style={{marginTop:10,display:'flex',alignItems:'center',gap:8}}>
                <BtnPrimary size="lg" disabled={!canRun} onClick={runAll}>
                  {running
                    ?<><Spin c={C.blue} s={14}/> Çalışıyor {progress&&`(${progress})`}</>
                    :selectedTools.length?`▶  ${selectedTools.length} Aracı Çalıştır`:'← Sol panelden araç seç'}
                </BtnPrimary>
                {running&&<span style={{fontSize:12,color:C.t3}}>{progress} tamamlandı</span>}
              </div>

              {error&&<div style={{marginTop:10,padding:'8px 12px',borderRadius:7,background:'rgba(239,68,68,0.08)',border:'1px solid rgba(239,68,68,0.2)',color:'#f87171',fontSize:12}}>⚠️ {error}</div>}
            </div>

            {/* Results */}
            <div className="qa-results" style={{flex:1,overflowY:'auto',padding:'16px 18px'}}>
              {results.length===0&&!running&&(
                <div style={{textAlign:'center',padding:'64px 0',color:C.t3}}>
                  <div style={{fontSize:44,marginBottom:12,opacity:.25}}>🧪</div>
                  <div style={{fontSize:14,fontWeight:700,color:C.t3,marginBottom:6}}>Sonuçlar burada görünecek</div>
                  <div style={{fontSize:12}}>Soldan araç seç → metni yaz → <span style={{color:C.blue,fontWeight:600}}>Çalıştır</span></div>
                </div>
              )}
              {running&&results.length===0&&(
                <div style={{textAlign:'center',padding:'64px 0'}}>
                  <Spin c={C.blue} s={28}/>
                  <div style={{marginTop:14,fontSize:13,color:C.blue}}>AI çalışıyor... {progress&&`${progress} tamamlandı`}</div>
                </div>
              )}
              {results.map((r,i)=><ResultCard key={i} tool={r.tool} content={r.content} messages={r.messages} onMessagesUpdate={msgs=>setResults(rs=>rs.map((x,j)=>j===i?{...x,messages:msgs}:x))} onClose={()=>setResults(rs=>rs.filter((_,j)=>j!==i))}/>)}
              {running&&results.length>0&&(
                <div style={{display:'flex',alignItems:'center',gap:8,color:C.blue,padding:'6px 0',fontSize:12}}><Spin c={C.blue} s={12}/> Devam ediyor... ({progress})</div>
              )}
            </div>
          </main>
        </div>
      )}

      {tab==='dashboard'&&<div style={{flex:1,overflow:'hidden',animation:tabTransition?'tabOut .18s ease forwards':'tabIn .18s ease'}}><DashboardTab usageStats={usageStats} allTools={allTools} bugs={bugs} setBugs={setBugs}/></div>}
      {tab==='manager'&&<div style={{flex:1,overflow:'hidden',animation:tabTransition?'tabOut .18s ease forwards':'tabIn .18s ease'}}><ToolManagerTab customTools={customTools} setCustomTools={setCT}/></div>}
      {editDefaultTool&&(
        <div style={{position:'fixed',inset:0,zIndex:9999,display:'flex',alignItems:'center',justifyContent:'center',background:'rgba(0,0,0,0.6)',backdropFilter:'blur(4px)'}}>
          <div style={{background:C.bg2,border:`1px solid ${C.b1}`,borderRadius:12,padding:'24px 28px',maxWidth:480,width:'92%',boxShadow:'0 20px 60px rgba(0,0,0,0.5)'}}>
            <div style={{fontSize:14,fontWeight:800,color:C.t0,marginBottom:16}}>✏️ Aracı Düzenle — {editDefaultTool.label}</div>
            <div style={{display:'flex',flexDirection:'column',gap:10,marginBottom:16}}>
              <Field label="Etiket"><TextInput value={editDefaultTool.label} onChange={e=>setEditDefaultTool(x=>({...x,label:e.target.value}))}/></Field>
              <Field label="Açıklama"><TextInput value={editDefaultTool.desc} onChange={e=>setEditDefaultTool(x=>({...x,desc:e.target.value}))}/></Field>
              <Field label="İkon">
                <div style={{display:'flex',gap:4,flexWrap:'wrap',maxHeight:80,overflowY:'auto'}}>
                  {['🔧','🧪','🐛','📄','✅','🔍','⚡','🎯','🔒','📊','🚀','🤖','🎭','🐍','📱','🌐','🔬','⚠️','📋','🗂','🔄','🎲','🧩','📝','🐞','📦','♿','🎨','🔁','📧','🔌','📮','🗄','🥒','❌','🛡️','🔐','📡','🏭','⚕️','⚙️'].map(ic=>(
                    <div key={ic} onClick={()=>setEditDefaultTool(x=>({...x,icon:ic}))} style={{width:30,height:30,borderRadius:6,display:'flex',alignItems:'center',justifyContent:'center',fontSize:16,cursor:'pointer',background:editDefaultTool.icon===ic?'rgba(59,158,237,0.2)':'rgba(255,255,255,0.03)',border:`1px solid ${editDefaultTool.icon===ic?C.blue:C.b0}`}}>{ic}</div>
                  ))}
                </div>
              </Field>
            </div>
            <div style={{display:'flex',gap:8}}>
              <BtnPrimary onClick={()=>{
                // localStorage'da editedTools kaydı tut
                const edited = LS.get('qa_edited_tools',{})
                edited[editDefaultTool.id] = editDefaultTool
                LS.set('qa_edited_tools', edited)
                setEditDefaultTool(null)
                // Sayfayı yenile ki değişiklik görünsün
                window.location.reload()
              }}>💾 Kaydet</BtnPrimary>
              <BtnGhost onClick={()=>setEditDefaultTool(null)}>İptal</BtnGhost>
            </div>
          </div>
        </div>
      )}
      {confirm&&<ConfirmDialog message={confirm.message} onConfirm={confirm.onConfirm} onCancel={()=>setConfirm(null)}/>}
      {tab==='assistant'&&<div style={{flex:1,overflow:'hidden',animation:tabTransition?'tabOut .18s ease forwards':'tabIn .18s ease'}}><AssistantTab/></div>}
      {tab==='process'&&<div style={{flex:1,overflow:'hidden',animation:tabTransition?'tabOut .18s ease forwards':'tabIn .18s ease'}}><ProcessTab/></div>}
      {tab==='defense'&&<div style={{flex:1,overflow:'hidden',animation:tabTransition?'tabOut .18s ease forwards':'tabIn .18s ease'}}><DefenseTab/></div>}
      {tab==='learning'&&<div style={{flex:1,overflow:'hidden',animation:tabTransition?'tabOut .18s ease forwards':'tabIn .18s ease'}}><LearningTab/></div>}
    </div>
      {/* Splash — app üzerinde, geçiş bitince unmount */}
      {!appVisible&&<SplashScreen exiting={splashOut} onEnter={enterApp}/>}
    </>
  )
}
