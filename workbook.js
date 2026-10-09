function safeOrderLink(value,fallback=''){try{const u=new URL(value);if(u.protocol==='https:'||u.protocol==='http:')return u.href;}catch{}return fallback;}
function readOrderWorkbook(buffer){
 const wb=XLSX.read(new Uint8Array(buffer),{type:'array',cellFormula:false,cellHTML:false});
 const ws=wb.Sheets['Parts list']||wb.Sheets[wb.SheetNames[0]];
 if(!ws||!ws['!ref'])throw Error('This workbook has no spreadsheet data.');
 const range=XLSX.utils.decode_range(ws['!ref']);if(range.e.r>10000)throw Error('This sheet is too large (maximum 10,000 rows).');
 const grid=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:true});
 const expected=['part name','sku','qty','cost per unit','total cost'];
 const headerIndex=grid.findIndex(r=>expected.every(h=>r.some(c=>String(c).trim().toLowerCase()===h)));
 if(headerIndex<0)throw Error('Expected columns: Part name, SKU, Qty, Cost per unit, Total cost.');
 const headers=grid[headerIndex].map(c=>String(c).trim().toLowerCase()),idx=expected.map(h=>headers.indexOf(h)),parsed=[];
 for(let i=headerIndex+1;i<grid.length;i++){
  const row=grid[i],sku=String(row[idx[1]]??'').trim(),name=String(row[idx[0]]??'').trim();
  if(!sku&&(!name||name.toLowerCase()==='total'))continue;
  const qty=Number(row[idx[2]]),price=Number(row[idx[3]]);
  if(!sku||!name||!valid(qty)||row[idx[3]]===''||!Number.isFinite(price)||price<0||price>1e9)throw Error(`Row ${i+range.s.r+1}: check the name, SKU, quantity, and unit cost.`);
  const product=catalog.find(p=>p.sku===sku);
  const cell=col=>ws[XLSX.utils.encode_cell({r:i+range.s.r,c:col+range.s.c})];
  const url=safeOrderLink(cell(idx[0])?.l?.Target,safeOrderLink(cell(idx[1])?.l?.Target,product?.url||''));
  if(!url)throw Error(`Row ${i+range.s.r+1}: no valid ordering link or matching catalog SKU.`);
  const existing=parsed.find(p=>p.sku===sku);
  if(existing){if(existing.price!==price)throw Error(`SKU ${sku} appears with different prices.`);if(!valid(existing.qty+qty))throw Error(`Quantity for ${sku} is too large.`);existing.qty+=qty;}
  else parsed.push({name,sku,qty,price,url});
 }
 if(!parsed.length)throw Error('The workbook contains no parts to resume.');
 return parsed;
}
async function previewWorkbook(){
 try{const wb=XLSX.read(new Uint8Array(await createXlsx(rows).arrayBuffer()),{type:'array'}),ws=wb.Sheets['Parts list'],grid=XLSX.utils.sheet_to_json(ws,{header:1,raw:true});
 $('workbook-preview').innerHTML='<table><thead><tr>'+grid[0].map(h=>`<th>${esc(h)}</th>`).join('')+'</tr></thead><tbody>'+grid.slice(1).map((r,i)=>`<tr class="${i===grid.length-2?'previewtotal':''}">`+Array.from({length:5},(_,j)=>{const cell=ws[XLSX.utils.encode_cell({r:i+1,c:j})],link=safeOrderLink(cell?.l?.Target),value=r[j]??'',display=j>=3&&typeof value==='number'?money(value):esc(value);return `<td>${link?`<a href="${esc(link)}" target="_blank" rel="noopener">${display}</a>`:display}</td>`;}).join('')+'</tr>').join('')+'</tbody></table>';
 $('workbook-dialog').showModal();
 }catch(e){$('status').textContent='Could not open the Excel preview. Please try again.';}
}
$('preview').addEventListener('click',previewWorkbook);
$('close-preview').addEventListener('click',()=>$('workbook-dialog').close());
$('preview-download').addEventListener('click',()=>downloadXlsx(rows));
$('upload').addEventListener('click',()=>$('upload-file').click());
$('upload-file').addEventListener('change',async e=>{
 const file=e.target.files[0];if(!file)return;
 try{if(file.size>10*1024*1024)throw Error('Choose an Excel file smaller than 10 MB.');
 const imported=readOrderWorkbook(await file.arrayBuffer());
 if(rows.length&&!confirm('Replace the current parts list with this workbook? Download the current list first if you want to keep it.'))return;
 rows=imported;render();$('status').textContent=`Resumed ${rows.length} parts from ${file.name}. Your list is ready to continue editing.`;
 }catch(error){$('status').textContent='Upload failed: '+error.message;}
 finally{e.target.value='';}
});

function sheetsClipboardData(items){
 const clean=value=>String(value).replace(/[\t\r\n]+/g,' ');
 const quote=value=>'"'+clean(value).replace(/"/g,'""')+'"';
 const link=(label,url)=>'=HYPERLINK('+quote(url)+','+quote(label)+')';
 const plain=items.map(r=>[link(r.name,r.url),link(r.sku,r.url),r.qty,r.price.toFixed(2),(cents(r)*r.qty/100).toFixed(2)].join('\t')).join('\n');
 const html='<html><body><table>'+items.map(r=>'<tr>'+[r.name,r.sku].map(value=>'<td><a href="'+esc(safeOrderLink(r.url))+'">'+esc(value)+'</a></td>').join('')+'<td>'+r.qty+'</td><td>'+r.price.toFixed(2)+'</td><td>'+(cents(r)*r.qty/100).toFixed(2)+'</td></tr>').join('')+'</table></body></html>';
 return {plain,html};
}
async function copyForSheets(){
 if(!rows.length)return;
 const {plain,html}=sheetsClipboardData(rows);
 try{
  let copied=false;
  if(navigator.clipboard?.write&&typeof ClipboardItem!=='undefined'){
   try{await navigator.clipboard.write([new ClipboardItem({'text/html':new Blob([html],{type:'text/html'}),'text/plain':new Blob([plain],{type:'text/plain'})})]);copied=true;}catch{}
  }
  if(!copied){
   if(navigator.clipboard?.writeText)await navigator.clipboard.writeText(plain);
   else{const input=document.createElement('textarea');input.value=plain;input.style.cssText='position:fixed;opacity:0';($('workbook-dialog').open?$('workbook-dialog'):document.body).append(input);input.select();const success=document.execCommand('copy');input.remove();if(!success)throw Error('Clipboard unavailable');}
  }
  const message=`Copied ${rows.length} part rows. Paste into Google Sheets with Ctrl+V (or ⌘V).`;
  $('copy-status').textContent=message;$('status').textContent=message;
 }catch{const message='Copy was blocked. Allow clipboard access in your browser and try again.';$('copy-status').textContent=message;$('status').textContent=message;}
}
$('copy-sheets').addEventListener('click',copyForSheets);
$('preview-copy').addEventListener('click',copyForSheets);
