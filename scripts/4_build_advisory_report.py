"""Build one advisory-labeled workbook from data/rows_<advisory>.json and cost data."""
import argparse, collections, datetime, json, os, re
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT=Path(__file__).resolve().parents[1]; DATA=ROOT/'data'
parser=argparse.ArgumentParser(); parser.add_argument('--advisory',required=True,choices=('PRFR-_4Z','JGW1-KG0')); args=parser.parse_args()
safe=args.advisory.replace('-','_')
profile=json.loads((ROOT/'scripts'/'advisories.json').read_text(encoding='utf-8'))[args.advisory]
rows=json.loads((DATA/f'rows_{safe}.json').read_text(encoding='utf-8'))
cost=json.loads((DATA/f'cost_{safe}.json').read_text(encoding='utf-8')) if (DATA/f'cost_{safe}.json').exists() else []
window=json.loads((DATA/f'cost_window_{safe}.json').read_text(encoding='utf-8')) if (DATA/f'cost_window_{safe}.json').exists() else {}
retire, price=rows['retirement_rows'], rows['price_rows']
location_map={'eu west':'westeurope','eu north':'northeurope','france central':'francecentral','france south':'francesouth','norway east':'norwayeast','norway west':'norwaywest','ap southeast':'southeastasia'}
regions=profile.get('regions',{})

groups=collections.defaultdict(float)
for item in cost:
    region=location_map.get((item.get('ResourceLocation') or '').lower(),(item.get('ResourceLocation') or '').lower())
    if profile['mode']=='prfr':
        meter=item.get('Meter') or ''
        label='v1' if any(x in meter for x in ('B','D','F','G','L','NP','HC')) and 'v2' not in meter else 'v2'
        rate=profile['priceIncreasePercent']/100
    else:
        label=region; rate=regions.get(region)
    if rate: groups[(label,rate)]+=item.get('Cost') or 0

fill=PatternFill('solid',fgColor='1F4E78'); white=Font(bold=True,color='FFFFFF',size=11); title=Font(bold=True,size=16,color='1F4E78'); sub=Font(bold=True,size=12,color='1F4E78'); note=Font(italic=True,size=9,color='555555'); bold=Font(bold=True); border=Border(*(Side(style='thin',color='D9D9D9') for _ in range(4))); money='#,##0.00'
def header(ws,row,n):
    for c in range(1,n+1): ws.cell(row,c).fill=fill;ws.cell(row,c).font=white;ws.cell(row,c).alignment=Alignment(horizontal='center',wrap_text=True);ws.cell(row,c).border=border
def table(ws,start,headers,data,widths,percent=()):
    for c,h in enumerate(headers,1):ws.cell(start,c,h)
    header(ws,start,len(headers))
    for r,item in enumerate(data,start+1):
        for c,h in enumerate(headers,1):
            cell=ws.cell(r,c,item.get(h,''));cell.border=border
            if c in percent and isinstance(cell.value,(int,float)):cell.number_format='0%'
    ws.freeze_panes=ws.cell(start+1,1);ws.auto_filter.ref=f'A{start}:{get_column_letter(len(headers))}{start+len(data)}'
    for c,w in enumerate(widths,1):ws.column_dimensions[get_column_letter(c)].width=w
    return start+len(data)+1
def total(ws,row,values):
    for c,v in enumerate(values,1):ws.cell(row,c,v).font=bold;ws.cell(row,c).border=border

wb=Workbook(); ws=wb.active;ws.title='Summary';ws.sheet_view.showGridLines=False
ws['A1']=f'{args.advisory} - {profile["title"]}';ws['A1'].font=title
ws['A2']=f'Generated {datetime.date.today().isoformat()} | VM-only assessment';ws['A2'].font=note
ws['A3']=f'Effective date: {profile["effectiveDate"]} | Cost window: {window.get("from","")[:10]} to {window.get("to","")[:10]}';ws['A3'].font=note
r=5
if profile['mode']=='prfr':
    ws.cell(r,1,'1 - Retirements by SKU Family').font=sub;r+=1
    counts=collections.Counter(x['Family'] for x in retire); fams=['Dv3','Dsv3','Ev3','Esv3']; data=[{'Family':f,'VM Count':counts.get(f,0)} for f in fams]
    r=table(ws,r,['Family','VM Count'],data,[20,16]);total(ws,r,['TOTAL',len(retire)]);r+=2
    ws.cell(r,1,'2 - Price Increase by SKU Family').font=sub;r+=1
    counts=collections.Counter((x['Impact Type'],x['Family']) for x in price); order=[('Price increase v1',f) for f in ['Bv1','D','Ds','F','Fs','G','Gs','Ls','NP','HC']]+[('Price increase v2',f) for f in ['Av2','Amv2','Dv2','Dsv2','Fsv2','Lsv2']]
    data=[{'Change Wave':i.replace('Price increase ',''),'Family':f,'VM Count':counts.get((i,f),0)} for i,f in order if counts.get((i,f),0) or any(k[0]==i for k in counts)]
    r=table(ws,r,['Change Wave','Family','VM Count'],data,[18,18,16]);total(ws,r,['TOTAL','',len(price)]);r+=2
else:
    ws.cell(r,1,'1 - Regional Price Increase by SKU Family').font=sub;r+=1
    counts=collections.Counter((x['Region'].lower(),x['Increase Rate'],x['Family']) for x in price);data=[{'Region':a,'Increase Rate':b,'Family':c,'VM Count':n} for (a,b,c),n in sorted(counts.items())]
    r=table(ws,r,['Region','Increase Rate','Family','VM Count'],data,[20,16,18,16],percent=(2,));total(ws,r,['','','TOTAL',len(price)]);r+=2
ws.cell(r,1,'2 - Overall Inventory Totals').font=sub;r+=1
r=table(ws,r,['Measure','Count'],[{'Measure':'Retirement VMs','Count':len(retire)},{'Measure':'Price-impacted VMs','Count':len(price)},{'Measure':'Price-impacted subscriptions','Count':len(rows['price_subs'])}],[32,16])+2
ws.cell(r,1,'3 - Estimated Price Impact - Amortized VM Cost').font=sub;r+=1
costrows=[]
for (label,rate),amount in sorted(groups.items()):costrows.append({'Region / Group':label.upper() if profile['mode']=='jgw' else label+' series','Increase Rate':rate,'Amortized Cost (30d)':round(amount,2),'Estimated Increase (30d)':round(amount*rate,2),'Projected Cost (30d)':round(amount*(1+rate),2)})
r=table(ws,r,['Region / Group','Increase Rate','Amortized Cost (30d)','Estimated Increase (30d)','Projected Cost (30d)'],costrows,[24,16,22,24,22],percent=(2,));total(ws,r,['TOTAL','',round(sum(x['Amortized Cost (30d)'] for x in costrows),2),round(sum(x['Estimated Increase (30d)'] for x in costrows),2),round(sum(x['Projected Cost (30d)'] for x in costrows),2)])
for note_text in profile['notes']:r+=1;ws.cell(r,1,'- '+note_text).font=note
headers=['Subscription ID','Subscription Name','Resource Group','VM Name','Region','VM Size','Family','Impact Type','Increase Rate','Power State','All Tags'];widths=[38,30,28,30,16,22,12,24,16,16,70]
if retire: table(wb.create_sheet('Retirement Inventory'),1,headers,sorted(retire,key=lambda x:x['VM Name']),widths,percent=(9,))
table(wb.create_sheet('Price Increase Inventory'),1,headers,sorted(price,key=lambda x:(x['Region'],x['VM Name'])),widths,percent=(9,))
cs=wb.create_sheet('Cost Impact');cs['A1']='Amortized VM cost used for estimate; RI-covered usage can overstate PAYG exposure.';cs['A1'].font=note
table(cs,3,['Region / Group','Increase Rate','Amortized Cost (30d)','Estimated Increase (30d)','Projected Cost (30d)'],costrows,[24,16,22,24,22],percent=(2,))
out=Path(os.environ.get('OUTNAME',ROOT/f'VM_Impact_{safe}.xlsx'));wb.save(out if out.is_absolute() else ROOT/out);print(f'Saved {out}')
