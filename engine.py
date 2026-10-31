
from __future__ import annotations
import math,statistics
from datetime import datetime,date,timedelta,timezone
import pandas as pd
TZ=timezone(timedelta(hours=7))
SCHEMA={
 'Settings':['key','value'],
 'Products':['sku','name','ng_rate','production_lead_days','transport_days'],
 'FG_Stock':['sku','on_hand','reserved_external','safety_target'],
 'Materials':['material_id','name','unit','lead_days','z','safety_stock_manual','moq','pack_size'],
 'RM_Stock':['material_id','on_hand','reserved_external'],
 'BOM':['sku','material_id','qty_per_start','startup_loss_per_lot'],
 'Demand_History':['sku','period','demand_qty'],
 'Orders':['order_id','customer','sku','qty','due_at','known_at'],
 'Purchase_Orders':['po_id','material_id','open_qty','expected_at','known_at'],
 'RM_Usage':['material_id','date','used_qty'],
 'Shipments':['shipment_id','order_id','accepted_qty','actual_arrival_at']}
def stamp(value):
 t=pd.Timestamp(value)
 if pd.isna(t):raise ValueError('วันเวลาไม่ถูกต้อง')
 if t.tzinfo is None:t=t.tz_localize(TZ)
 return t.to_pydatetime().astimezone(TZ)
def iso(t):return t.isoformat(timespec='seconds') if t is not None else None
def period_add(period,n):
 y,m=map(int,period.split('-'));y,m=divmod(y*12+m-1+n,12);return f'{y:04d}-{m+1:02d}'
def end_month(period):return stamp(period_add(period,1)+'-01T00:00:00+07:00')-timedelta(seconds=1)
def q_start(g,r):
 if not math.isfinite(g) or g<0 or not 0<=r<1:raise ValueError('Good ต้องไม่ติดลบ และ NG ต้องอยู่ [0,1)')
 return 0 if g==0 else math.ceil(g/(1-r)-1e-10)
def number(value,label,integer=False,nullable=False):
 if value is None or pd.isna(value) or str(value).strip()=='':
  if nullable:return None
  raise ValueError(label+': ห้ามว่าง')
 try:x=float(value)
 except (ValueError,TypeError):raise ValueError(label+': ต้องเป็นตัวเลข')
 if not math.isfinite(x) or x<0 or (integer and not x.is_integer()):raise ValueError(label+': ต้องไม่ติดลบ'+('และเป็นจำนวนเต็ม' if integer else ''))
 return int(x) if integer else x
def read_workbook(source):
 tabs=pd.read_excel(source,sheet_name=None,engine='openpyxl',dtype=object)
 missing=[s for s in SCHEMA if s not in tabs]
 if missing:raise ValueError('ไม่พบชีต: '+', '.join(missing))
 for s,columns in SCHEMA.items():
  absent=[c for c in columns if c not in tabs[s].columns]
  if absent:raise ValueError(s+' ขาดคอลัมน์ '+', '.join(absent))
  if len(tabs[s])>10000:raise ValueError(s+': มากกว่า10,000แถว เกินขอบเขตต้นแบบ')
 data={s:tabs[s][cols].dropna(how='all').where(pd.notna(tabs[s][cols]),None).to_dict('records') for s,cols in SCHEMA.items()}
 numeric={'Products':{'ng_rate':False,'production_lead_days':False,'transport_days':False},'FG_Stock':{'on_hand':True,'reserved_external':True,'safety_target':True},'Materials':{'lead_days':False,'z':False,'safety_stock_manual':False,'moq':False,'pack_size':False},'RM_Stock':{'on_hand':False,'reserved_external':False},'BOM':{'qty_per_start':False,'startup_loss_per_lot':False},'Demand_History':{'demand_qty':True},'Orders':{'qty':True},'Purchase_Orders':{'open_qty':False},'RM_Usage':{'used_qty':False},'Shipments':{'accepted_qty':True}}
 times={'Orders':['due_at','known_at'],'Purchase_Orders':['expected_at','known_at'],'Shipments':['actual_arrival_at']}
 for s,rows in data.items():
  if s=='Settings':continue
  for i,row in enumerate(rows,2):
   for k in SCHEMA[s]:
    if k in numeric.get(s,{}):row[k]=number(row[k],f'{s}!{k} แถว{i}',numeric[s][k],nullable=(s=='Materials' and k=='safety_stock_manual'))
    elif k in times.get(s,[]):row[k]=iso(stamp(row[k]))
    elif k=='date':row[k]=stamp(row[k]).date().isoformat()
    elif k=='period':
     v=row[k];row[k]=v.strftime('%Y-%m') if hasattr(v,'strftime') else str(v).strip()
    else:
     if row[k] is None or not str(row[k]).strip():raise ValueError(f'{s}!{k} แถว{i}: ห้ามว่าง')
     row[k]=str(row[k]).strip()
 settings={str(r['key']).strip():r['value'] for r in data['Settings'] if r['key'] is not None}
 if len(settings)!=len(data['Settings']):raise ValueError('Settings key ว่าง/ซ้ำ')
 for k in ['as_of','horizon_days','include_forecast','data_label']:
  if k not in settings:raise ValueError('Settings ขาด '+k)
 settings['as_of']=iso(stamp(settings['as_of']));settings['horizon_days']=number(settings['horizon_days'],'horizon_days',True)
 if not 1<=settings['horizon_days']<=180:raise ValueError('horizon_days ต้องอยู่1–180')
 value=str(settings['include_forecast']).strip().lower()
 if value not in ('true','false','1','0'):raise ValueError('include_forecast ให้ใช้ TRUE/FALSE')
 settings['include_forecast']=value in ('true','1');settings['data_label']=str(settings['data_label']);data['Settings']=settings
 validate(data);return data,tabs

def validate(d):
 asof=stamp(d['Settings']['as_of'])
 def unique(sheet,keys):
  seen=set()
  for row in d[sheet]:
   value=tuple(row[k] for k in keys)
   if value in seen:raise ValueError(sheet+': รหัสซ้ำ '+str(value))
   seen.add(value)
 for s,keys in [('Products',['sku']),('FG_Stock',['sku']),('Materials',['material_id']),('RM_Stock',['material_id']),('BOM',['sku','material_id']),('Demand_History',['sku','period']),('Orders',['order_id']),('Purchase_Orders',['po_id']),('RM_Usage',['material_id','date']),('Shipments',['shipment_id'])]:unique(s,keys)
 ps={x['sku'] for x in d['Products']};ms={x['material_id'] for x in d['Materials']};orders={x['order_id']:x for x in d['Orders']}
 if not ps or not ms:raise ValueError('ต้องมีสินค้าและวัสดุอย่างน้อยอย่างละ1')
 if {x['sku'] for x in d['FG_Stock']}!=ps:raise ValueError('FG_Stock ต้องมี1แถวต่อสินค้า และรหัสตรง Products')
 if {x['material_id'] for x in d['RM_Stock']}!=ms:raise ValueError('RM_Stock ต้องมี1แถวต่อวัสดุ และรหัสตรง Materials')
 for p in d['Products']:
  if not 0<=p['ng_rate']<1:raise ValueError(p['sku']+': ng_rate ต้องต่ำกว่า1')
  if p['production_lead_days']>90 or p['transport_days']>90:raise ValueError('Lead time สินค้าเกิน90วัน')
 for m in d['Materials']:
  if m['pack_size']<=0 or m['z']>5 or m['lead_days']>365:raise ValueError('pack_size/z/lead_days ผิดขอบเขต')
 for s in ['FG_Stock','RM_Stock']:
  for x in d[s]:
   if x['reserved_external']>x['on_hand']:raise ValueError(s+': จองมากกว่าคงเหลือ')
 for b in d['BOM']:
  if b['sku'] not in ps or b['material_id'] not in ms or b['qty_per_start']<=0:raise ValueError('BOM รหัส/อัตราใช้ไม่ถูกต้อง')
 for h in d['Demand_History']:
  if h['sku'] not in ps:raise ValueError('Demand_History อ้างสินค้าไม่พบ')
  try:
   if len(h['period'])!=7 or period_add(h['period'],0)!=h['period']:raise ValueError()
  except Exception:raise ValueError('Demand period ต้องเป็น YYYY-MM')
 for c in d['RM_Usage']:
  if c['material_id'] not in ms:raise ValueError('RM_Usage อ้างวัสดุไม่พบ')
 for o in d['Orders']:
  if o['sku'] not in ps or o['qty']<=0:raise ValueError('Orders SKU/จำนวนผิด')
  if stamp(o['known_at'])>asof or stamp(o['due_at'])<stamp(o['known_at']):raise ValueError('Orders วันรับทราบ/กำหนดไม่สอดคล้องกับ as_of')
 for r in d['Purchase_Orders']:
  if r['material_id'] not in ms or stamp(r['known_at'])>asof:raise ValueError('PO รหัส/วันทราบผิด')
 totals={k:0 for k in orders}
 for s in d['Shipments']:
  if s['order_id'] not in orders or s['accepted_qty']<=0 or stamp(s['actual_arrival_at'])>asof:raise ValueError('Shipments รหัส/จำนวน/วันรับมอบไม่ถูกต้อง')
  totals[s['order_id']]+=s['accepted_qty']
 for k,v in totals.items():
  if v>orders[k]['qty']:raise ValueError('ส่งเกินยอด Order '+k)

def forecast(history,sku,target):
 h=sorted([x for x in history if x['sku']==sku and x['period']<target],key=lambda x:x['period'])
 if len(h)<4 or h[-1]['period']!=period_add(target,-1):return {'sku':sku,'error':'ต้องมีอย่างน้อย4เดือนและจบเดือนก่อนเดือนพยากรณ์','methods':[]}
 if any(h[i]['period']!=period_add(h[i-1]['period'],1) for i in range(1,len(h))):return {'sku':sku,'error':'ประวัติรายเดือนขาดช่วง','methods':[]}
 a=[x['demand_qty'] for x in h];ses={};prev=a[0]
 for i in range(1,len(a)+1):prev=.3*a[i-1]+.7*prev;ses[i]=prev
 def calc(name,i):return sum(a[i-3:i])/3 if name=='SMA' else .5*a[i-1]+.3*a[i-2]+.2*a[i-3] if name=='WMA' else ses[i]
 methods=[]
 for name in ['SMA','WMA','SES']:
  errors=[{'period':h[i]['period'],'actual':a[i],'forecast':calc(name,i)} for i in range(3,len(a))];n=len(errors);valid=[e for e in errors if e['actual']>0];mad=sum(abs(e['forecast']-e['actual']) for e in errors)/n;mse=sum((e['forecast']-e['actual'])**2 for e in errors)/n
  methods.append({'method':name,'forecast':calc(name,len(a)),'MAD':mad,'MSE':mse,'RMSE':math.sqrt(mse),'MAPE':100*sum(abs(e['forecast']-e['actual'])/e['actual'] for e in valid)/len(valid) if valid else None,'n':n,'n_valid':len(valid),'errors':errors})
 selected=min(methods,key=lambda x:(x['MAD'],x['MSE'],['SMA','WMA','SES'].index(x['method'])))
 return {'sku':sku,'target':target,'methods':methods,'selected':selected,'history':h}
def delivery(order,shipments,asof):
 ev=sorted([s for s in shipments if s['order_id']==order['order_id'] and stamp(s['actual_arrival_at'])<=asof],key=lambda s:stamp(s['actual_arrival_at']));total=0;ontime=0;full=None;due=stamp(order['due_at'])
 for s in ev:
  total+=s['accepted_qty']
  if stamp(s['actual_arrival_at'])<=due:ontime+=s['accepted_qty']
  if total>=order['qty'] and full is None:full=stamp(s['actual_arrival_at'])
 return {'delivered':min(total,order['qty']),'on_time_qty':min(ontime,order['qty']),'backlog':max(0,order['qty']-total),'full_at':iso(full),'late_minutes':max(0,(full-due).total_seconds()/60) if full else None,'otd':int(ontime>=order['qty']) if due<=asof else None,'status':('COMPLETE_ON_TIME' if full<=due else 'COMPLETE_LATE') if full else ('OPEN_LATE' if asof>=due else 'OPEN')}

def calculate(d):
 validate(d);settings=d['Settings'];asof=stamp(settings['as_of']);target=asof.strftime('%Y-%m');end=asof+timedelta(days=settings['horizon_days']);pmap={x['sku']:x for x in d['Products']};fg={x['sku']:dict(x) for x in d['FG_Stock']};fgfree={k:x['on_hand']-x['reserved_external'] for k,x in fg.items()};stocks={x['material_id']:x for x in d['RM_Stock']};warnings=[];orders=[];jobs=[];needs=[]
 forecasts=[forecast(d['Demand_History'],sku,target) for sku in pmap]
 for o in sorted(d['Orders'],key=lambda x:(stamp(x['due_at']),x['order_id'])):
  pr=pmap[o['sku']];de=delivery(o,d['Shipments'],asof);allocated=min(de['backlog'],fgfree[o['sku']]);fgfree[o['sku']]-=allocated;g=de['backlog']-allocated;req=stamp(o['due_at'])-timedelta(days=pr['production_lead_days']+pr['transport_days']);use=max(asof,req);q=q_start(g,pr['ng_rate']);row={**o,**de,'fg_allocated_plan':allocated,'good_to_make':g,'start_qty':q,'material_required_at':iso(req),'planned_use_at':iso(use),'in_horizon':use<=end,'late_plan':g>0 and req<asof,'material_risks':[]};orders.append(row)
  if g>0 and use<=end:jobs.append({'job_id':o['order_id'],'order_id':o['order_id'],'sku':o['sku'],'kind':'FIRM','good':g,'start_qty':q,'use_at':iso(use)})
 for f in forecasts:
  pr=pmap[f['sku']];firm=sum(o['qty'] for o in d['Orders'] if o['sku']==pr['sku'] and stamp(o['due_at']).strftime('%Y-%m')==target)
  if f.get('error'):warnings.append(pr['sku']+': '+f['error'])
  residual=max(0,f['selected']['forecast']-firm) if settings['include_forecast'] and not f.get('error') else 0;desired=max(residual,fg[pr['sku']]['safety_target']);covered=min(desired,fgfree[pr['sku']]);fgfree[pr['sku']]-=covered;g=math.ceil(max(0,desired-covered)-1e-10);f.update(firm_month=firm,forecast_residual=residual,buffer_good=g)
  use=max(asof,end_month(target)-timedelta(days=pr['production_lead_days']+pr['transport_days']))
  if g>0 and use<=end:jobs.append({'job_id':'BUFFER:'+pr['sku'],'order_id':None,'sku':pr['sku'],'kind':'BUFFER','good':g,'start_qty':q_start(g,pr['ng_rate']),'use_at':iso(use)})
 omap={o['order_id']:o for o in orders}
 for j in jobs:
  rows=[b for b in d['BOM'] if b['sku']==j['sku']]
  if not rows:
   warnings.append(j['job_id']+': ไม่มี BOM จึงยืนยันความพร้อมวัสดุไม่ได้')
   if j['order_id']:omap[j['order_id']]['material_risks'].append('ไม่มี BOM')
  for b in rows:
   run=j['start_qty']*b['qty_per_start'];needs.append({**j,'material_id':b['material_id'],'bom_per_start':b['qty_per_start'],'run_material':round(run,6),'startup_loss':b['startup_loss_per_lot'],'gross_requirement':round(run+b['startup_loss_per_lot'],6)})
 materials=[];timeline=[]
 for m in d['Materials']:
  mid=m['material_id'];stock=stocks[mid];free=stock['on_hand']-stock['reserved_external'];sample=sorted([c for c in d['RM_Usage'] if c['material_id']==mid and c['date']<asof.date().isoformat()],key=lambda x:x['date'])[-60:];continuous=all((date.fromisoformat(sample[i]['date'])-date.fromisoformat(sample[i-1]['date'])).days==1 for i in range(1,len(sample)));vals=[x['used_qty'] for x in sample];mean=statistics.mean(vals) if vals and continuous else None;sd=statistics.stdev(vals) if len(vals)>=2 and continuous else None;ss=m['safety_stock_manual'] if m['safety_stock_manual'] is not None else m['z']*sd*math.sqrt(m['lead_days']) if sd is not None else None;rop=mean*m['lead_days']+ss if mean is not None and ss is not None else None
  receipts=[r for r in d['Purchase_Orders'] if r['material_id']==mid and r['open_qty']>0];overdue=[r for r in receipts if stamp(r['expected_at'])<asof];valid=[r for r in receipts if asof<=stamp(r['expected_at'])<=end];uses=[n for n in needs if n['material_id']==mid];events=[{'at':r['expected_at'],'kind':'RECEIPT','reference':r['po_id'],'receipt_qty':r['open_qty'],'need_qty':0,'order_id':None} for r in valid]+[{'at':n['use_at'],'kind':n['kind'],'reference':n['job_id'],'receipt_qty':0,'need_qty':n['gross_requirement'],'order_id':n['order_id']} for n in uses];events.sort(key=lambda e:(stamp(e['at']),0 if e['kind']=='RECEIPT' else 1,e['reference']))
  bal=free;minimum=free;first_short=None;first_zero=asof if free<=0 else None;first_ss=asof if ss is not None and free<ss else None
  timeline.append({'material_id':mid,'at':iso(asof),'kind':'SNAPSHOT','reference':'opening','receipt_qty':0,'need_qty':0,'projected_balance':bal})
  for e in events:
   bal=round(bal+e['receipt_qty']-e['need_qty'],6);minimum=min(minimum,bal);at=stamp(e['at'])
   if e['need_qty']>0 and bal<0 and first_short is None:first_short=at
   if e['need_qty']>0 and bal<=0 and first_zero is None:first_zero=at
   if ss is not None and bal<ss and first_ss is None:first_ss=at
   if e['order_id'] and bal<0:omap[e['order_id']]['material_risks'].append(mid+f' ขาดสะสม {-bal:.6g} '+m['unit'])
   timeline.append({'material_id':mid,**{k:v for k,v in e.items() if k!='order_id'},'projected_balance':bal})
  ip=free+sum(r['open_qty'] for r in receipts);nr=max(0,(ss or 0)-minimum);suggest=max(m['moq'],math.ceil(nr/m['pack_size']-1e-10)*m['pack_size']) if nr>0 else 0;critical=first_ss or first_short;latest=critical-timedelta(days=m['lead_days']) if critical else None
  status='SHORT' if first_short else 'DATA' if ss is None or mean is None else 'LOW' if first_ss else 'REVIEW' if ip<=rop else 'OK'
  if overdue:warnings.append(mid+': PO เลยคาดรับ ยังไม่ถือว่ารับแล้ว: '+', '.join(r['po_id'] for r in overdue))
  if mean is None or ss is None:warnings.append(mid+': ประวัติใช้รายวัน/SS ไม่ครบ ตรวจวันใช้0และวันที่ขาด')
  materials.append({**m,'on_hand':stock['on_hand'],'reserved_external':stock['reserved_external'],'free_stock':free,'daily_mean':mean,'daily_sd':sd,'sample_n':len(sample),'SS':ss,'ROP':rop,'IP':ip,'days_cover':free/mean if mean else None,'gross_requirement':sum(n['gross_requirement'] for n in uses),'minimum_projected_stock':minimum,'projected_end_stock':bal,'first_below_SS':iso(first_ss),'first_zero':iso(first_zero),'first_shortage':iso(first_short),'latest_order_at':iso(latest),'normal_order_late':bool(latest and latest<asof),'suggest_order_qty':suggest,'status':status})
 for o in orders:o['readiness']='CLOSED' if o['backlog']==0 else 'FG_READY' if o['fg_allocated_plan']>=o['backlog'] else 'OUTSIDE' if not o['in_horizon'] else 'MATERIAL_RISK' if o['material_risks'] else 'LATE_PLAN' if o['late_plan'] else 'MATERIAL_COVERED'
 cohort=[o for o in orders if stamp(o['due_at'])<=asof];qty=sum(o['qty'] for o in cohort)
 kpi={'OTD_calloff':100*sum(o['otd'] for o in cohort)/len(cohort) if cohort else None,'on_time_quantity':100*sum(o['on_time_qty'] for o in cohort)/qty if qty else None,'cumulative_fulfilment':100*sum(o['delivered'] for o in cohort)/qty if qty else None,'due_count':len(cohort),'backlog_all':sum(o['backlog'] for o in orders)}
 return {'settings':settings,'target_month':target,'forecasts':forecasts,'orders':orders,'materials':materials,'jobs':jobs,'requirements':needs,'timeline':timeline,'warnings':warnings,'kpi':kpi,'fg_remaining_plan':fgfree}
