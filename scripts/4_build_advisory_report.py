"""Build one advisory-labeled workbook from data/rows_<advisory>.json and cost data."""
import argparse, collections, datetime, json, os, re
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import classify as cls

ROOT=Path(__file__).resolve().parents[1]; DATA=ROOT/'data'
parser=argparse.ArgumentParser(); parser.add_argument('--advisory',required=True,choices=('PRFR-_4Z','JGW1-KG0')); args=parser.parse_args()
safe=args.advisory.replace('-','_')
config=json.loads((ROOT/'config.json').read_text(encoding='utf-8')) if (ROOT/'config.json').exists() else {}
org=(config.get('organizationName') or '').strip() or 'Azure'
org_slug=re.sub(r'[^A-Za-z0-9]+','_',org).strip('_') or 'Azure'
profile=json.loads((ROOT/'scripts'/'advisories.json').read_text(encoding='utf-8'))[args.advisory]
rows=json.loads((DATA/f'rows_{safe}.json').read_text(encoding='utf-8'))
cost=json.loads((DATA/f'cost_{safe}.json').read_text(encoding='utf-8')) if (DATA/f'cost_{safe}.json').exists() else []
storage_cost=json.loads((DATA/f'storage_cost_{safe}.json').read_text(encoding='utf-8')) if (DATA/f'storage_cost_{safe}.json').exists() else []
cost_errors=json.loads((DATA/f'cost_errors_{safe}.json').read_text(encoding='utf-8')) if (DATA/f'cost_errors_{safe}.json').exists() else []
window=json.loads((DATA/f'cost_window_{safe}.json').read_text(encoding='utf-8')) if (DATA/f'cost_window_{safe}.json').exists() else {}
submap=json.loads((DATA/'submap.json').read_text(encoding='utf-8')) if (DATA/'submap.json').exists() else {}
retire, price=rows['retirement_rows'], rows['price_rows']
location_map={'eu west':'westeurope','eu north':'northeurope','france central':'francecentral','france south':'francesouth','norway east':'norwayeast','norway west':'norwaywest','ap southeast':'southeastasia'}
regions=profile.get('regions',{})
excluded_storage=[term.lower() for term in profile.get('excludedStorageTerms', [])]

groups=collections.defaultdict(float)
for item in cost:
    if str(item.get('PricingModel') or '') == 'Reservation':
        continue
    if item.get('MeterCategory') != 'Virtual Machines':
        continue
    region=location_map.get((item.get('ResourceLocation') or '').lower(),(item.get('ResourceLocation') or '').lower())
    if profile['mode']=='prfr':
        meter=item.get('Meter') or ''
        label, wave = cls.meter_group(meter)
        if not label:
            continue
        label = wave
        rate=profile['priceIncreasePercent']/100
    else:
        label=region; rate=regions.get(region)
    if rate: groups[(label,rate)]+=item.get('Cost') or 0

storage_groups=collections.defaultdict(float)
if profile['mode']=='jgw':
    for item in storage_cost:
        region=location_map.get((item.get('ResourceLocation') or '').lower(),(item.get('ResourceLocation') or '').lower())
        meter_text=' '.join(str(item.get(key) or '') for key in ('MeterSubCategory','Meter')).lower()
        if str(item.get('PricingModel') or '') != 'Reservation' and region in regions and not any(term in meter_text for term in excluded_storage):
            storage_groups[(region, regions[region], item.get('MeterSubCategory') or item.get('Meter') or 'Unknown')] += item.get('Cost') or 0

fill=PatternFill('solid',fgColor='1F4E78'); white=Font(bold=True,color='FFFFFF',size=11); title=Font(bold=True,size=16,color='1F4E78'); sub=Font(bold=True,size=12,color='1F4E78'); note=Font(italic=True,size=9,color='555555'); bold=Font(bold=True); border=Border(*(Side(style='thin',color='D9D9D9') for _ in range(4))); money='#,##0.00'
def header(ws,row,n):
    for c in range(1,n+1): ws.cell(row,c).fill=fill;ws.cell(row,c).font=white;ws.cell(row,c).alignment=Alignment(horizontal='center',wrap_text=True);ws.cell(row,c).border=border
def table(ws,start,headers,data,widths,percent=(),freeze=True,autofilter=True):
    for c,h in enumerate(headers,1):ws.cell(start,c,h)
    header(ws,start,len(headers))
    for r,item in enumerate(data,start+1):
        for c,h in enumerate(headers,1):
            cell=ws.cell(r,c,item.get(h,''));cell.border=border
            if c in percent and isinstance(cell.value,(int,float)):cell.number_format='0%'
    if freeze:ws.freeze_panes=ws.cell(start+1,1)
    if autofilter:ws.auto_filter.ref=f'A{start}:{get_column_letter(len(headers))}{start+len(data)}'
    for c,w in enumerate(widths,1):ws.column_dimensions[get_column_letter(c)].width=w
    return start+len(data)+1
def total(ws,row,values):
    for c,v in enumerate(values,1):ws.cell(row,c,v).font=bold;ws.cell(row,c).border=border

wb=Workbook(); ws=wb.active;ws.title='Summary';ws.sheet_view.showGridLines=False
ws['A1']=f'{org} - {args.advisory} - {profile["title"]}';ws['A1'].font=title
scope_label = 'VM and Storage assessment' if profile['mode'] == 'jgw' else 'VM assessment'
ws['A2']=f'Generated {datetime.date.today().isoformat()} | {scope_label} | tenant-wide cost scope';ws['A2'].font=note
ws['A3']=f'Effective date: {profile["effectiveDate"]} | Cost window: {window.get("from","")[:10]} to {window.get("to","")[:10]}';ws['A3'].font=note
if cost_errors:
    ws['A4']=f'WARNING: cost query failed for {len(cost_errors)} tenant subscriptions; cost totals are partial. See the "Failed Subscriptions" tab.'
    ws['A4'].font=Font(bold=True,color='C00000')
r=5
if profile['mode']=='prfr':
    ws.cell(r,1,'1 - Retirements by SKU Family').font=sub;r+=1
    counts=collections.Counter(x['Family'] for x in retire); fams=['Dv3','Dsv3','Ev3','Esv3']; data=[{'Family':f,'VM Count':counts.get(f,0)} for f in fams]
    r=table(ws,r,['Family','VM Count'],data,[20,16],freeze=False,autofilter=False);total(ws,r,['TOTAL',len(retire)]);r+=2
    ws.cell(r,1,'2 - Price Increase by SKU Family').font=sub;r+=1
    counts=collections.Counter((x['Impact Type'],x['Family']) for x in price); order=[('Price increase v1',f) for f in ['Bv1','D','Ds','F','Fs','G','Gs','Ls','NP','HC']]+[('Price increase v2',f) for f in ['Av2','Amv2','Dv2','Dsv2','Fsv2','Lsv2']]
    data=[{'Change Wave':i.replace('Price increase ',''),'Family':f,'VM Count':counts.get((i,f),0)} for i,f in order if counts.get((i,f),0) or any(k[0]==i for k in counts)]
    r=table(ws,r,['Change Wave','Family','VM Count'],data,[18,18,16],freeze=False,autofilter=False);total(ws,r,['TOTAL','',len(price)]);r+=2
else:
    ws.cell(r,1,'1 - Regional Price Increase by SKU Family').font=sub;r+=1
    counts=collections.Counter((x['Region'].lower(),x['Increase Rate'],x['Family']) for x in price);data=[{'Region':a,'Increase Rate':b,'Family':c,'VM Count':n} for (a,b,c),n in sorted(counts.items())]
    r=table(ws,r,['Region','Increase Rate','Family','VM Count'],data,[20,16,18,16],percent=(2,),freeze=False,autofilter=False);total(ws,r,['','','TOTAL',len(price)]);r+=2
ws.cell(r,1,'2 - Overall Inventory Totals').font=sub;r+=1
r=table(ws,r,['Measure','Count'],[{'Measure':'Retirement VMs','Count':len(retire)},{'Measure':'Price-impacted VMs','Count':len(price)},{'Measure':'Price-impacted subscriptions','Count':len(rows['price_subs'])}],[32,16],freeze=False,autofilter=False)+2
ws.cell(r,1,'3 - Estimated Price Impact - Amortized VM Cost').font=sub;r+=1
costrows=[]
for (label,rate),amount in sorted(groups.items()):costrows.append({'Region / Group':label.upper() if profile['mode']=='jgw' else label+' series','Increase Rate':rate,'Amortized Cost (30d)':round(amount,2),'Estimated Increase (30d)':round(amount*rate,2),'Projected Cost (30d)':round(amount*(1+rate),2)})
r=table(ws,r,['Region / Group','Increase Rate','Amortized Cost (30d)','Estimated Increase (30d)','Projected Cost (30d)'],costrows,[24,16,22,24,22],percent=(2,),freeze=False,autofilter=False);total(ws,r,['TOTAL','',round(sum(x['Amortized Cost (30d)'] for x in costrows),2),round(sum(x['Estimated Increase (30d)'] for x in costrows),2),round(sum(x['Projected Cost (30d)'] for x in costrows),2)])
for note_text in profile['notes']:r+=1;ws.cell(r,1,'- '+note_text).font=note
if profile['mode']=='jgw':
    r += 2
    ws.cell(r,1,'4 - Estimated Storage Price Impact - Amortized Storage Cost').font=sub
    r += 1
    storage_rows=[]
    for (region, rate, meter), amount in sorted(storage_groups.items()):
        storage_rows.append({'Region':region.upper(),'Increase Rate':rate,'Storage Meter Group':meter,'Amortized Cost (30d)':round(amount,2),'Estimated Increase (30d)':round(amount*rate,2),'Projected Cost (30d)':round(amount*(1+rate),2)})
    r=table(ws,r,['Region','Increase Rate','Storage Meter Group','Amortized Cost (30d)','Estimated Increase (30d)','Projected Cost (30d)'],storage_rows,[20,16,32,22,24,22],percent=(2,),freeze=False,autofilter=False)
    total(ws,r,['TOTAL','', '',round(sum(x['Amortized Cost (30d)'] for x in storage_rows),2),round(sum(x['Estimated Increase (30d)'] for x in storage_rows),2),round(sum(x['Projected Cost (30d)'] for x in storage_rows),2)])
    r += 2
    ws.cell(r,1,'5 - Combined Cost Impact (VM + Storage)').font=sub
    r += 1
    vm_amort=sum(x['Amortized Cost (30d)'] for x in costrows); vm_inc=sum(x['Estimated Increase (30d)'] for x in costrows); vm_proj=sum(x['Projected Cost (30d)'] for x in costrows)
    st_amort=sum(x['Amortized Cost (30d)'] for x in storage_rows); st_inc=sum(x['Estimated Increase (30d)'] for x in storage_rows); st_proj=sum(x['Projected Cost (30d)'] for x in storage_rows)
    combo=[{'Category':'VM','Amortized Cost (30d)':round(vm_amort,2),'Estimated Increase (30d)':round(vm_inc,2),'Projected Cost (30d)':round(vm_proj,2)},
           {'Category':'Storage','Amortized Cost (30d)':round(st_amort,2),'Estimated Increase (30d)':round(st_inc,2),'Projected Cost (30d)':round(st_proj,2)}]
    r=table(ws,r,['Category','Amortized Cost (30d)','Estimated Increase (30d)','Projected Cost (30d)'],combo,[20,22,24,22],freeze=False,autofilter=False)
    total(ws,r,['TOTAL',round(vm_amort+st_amort,2),round(vm_inc+st_inc,2),round(vm_proj+st_proj,2)])
headers=['Subscription ID','Subscription Name','Resource Group','VM Name','Region','VM Size','Family','Impact Type','Increase Rate','Power State','All Tags'];widths=[38,30,28,30,16,22,12,24,16,16,70]
if retire: table(wb.create_sheet('Retirement Inventory'),1,headers,sorted(retire,key=lambda x:x['VM Name']),widths,percent=(9,))
table(wb.create_sheet('Price Increase Inventory'),1,headers,sorted(price,key=lambda x:(x['Region'],x['VM Name'])),widths,percent=(9,))
cs=wb.create_sheet('VM Cost Impact');cs['A1']='Tenant-wide Amortized VM cost excluding PricingModel = Reservation; SavingsPlan usage remains included.';cs['A1'].font=note
table(cs,3,['Region / Group','Increase Rate','Amortized Cost (30d)','Estimated Increase (30d)','Projected Cost (30d)'],costrows,[24,16,22,24,22],percent=(2,))
if profile['mode']=='jgw':
    storage_sheet=wb.create_sheet('Storage Cost Impact')
    storage_sheet['A1']='Amortized Storage cost in JGW1-KG0 affected regions, excluding notice-listed Storage services.';storage_sheet['A1'].font=note
    table(storage_sheet,3,['Region','Increase Rate','Storage Meter Group','Amortized Cost (30d)','Estimated Increase (30d)','Projected Cost (30d)'],storage_rows,[20,16,32,22,24,22],percent=(2,))
def failure_reason(code):
    if code==429:return 'Throttled (HTTP 429) - per-subscription Cost Management rate limit, often caused by the subscription\'s own cost tooling. Not included in totals; re-run later to capture it.'
    if code in (401,403):return f'Authorization failed (HTTP {code}) - the running identity lacks Cost Management Reader on this subscription.'
    if code and 500<=int(code)<=599:return f'Azure service error (HTTP {code}).'
    if not code:return 'Network/transport error (no HTTP status returned).'
    return f'Query failed (HTTP {code}).'
fs=wb.create_sheet('Failed Subscriptions')
fs.sheet_view.showGridLines=False
fs['A1']=f'{len(cost_errors)} subscription(s) did not return cost data after all retries and recovery passes. Their spend is NOT included in the cost totals in this workbook.';fs['A1'].font=sub
fs['A2']='Re-run the cost step (optionally with more recovery passes) to capture transient throttles. Persistent failures usually indicate the subscription rate-limits Cost Management queries or the identity lacks access.';fs['A2'].font=note
if cost_errors:
    err_data=[{'Subscription ID':e.get('sub'),'Subscription Name':submap.get(e.get('sub'),'(unknown)'),'HTTP Code':e.get('code'),'Reason':failure_reason(e.get('code'))} for e in cost_errors]
    table(fs,4,['Subscription ID','Subscription Name','HTTP Code','Reason'],err_data,[38,34,12,90])
else:
    fs['A4']='None - all tenant subscriptions returned cost data.';fs['A4'].font=bold
out=Path(os.environ.get('OUTNAME',ROOT/f'{org_slug}_VM_Impact_{safe}.xlsx'));wb.save(out if out.is_absolute() else ROOT/out);print(f'Saved {out}')
