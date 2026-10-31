
from pathlib import Path
from io import BytesIO
from html import escape
import copy,hashlib,math
import pandas as pd
import streamlit as st
from engine import read_workbook,calculate,stamp

st.set_page_config(page_title='GateCAT | วัสดุพร้อม ส่งทัน',page_icon='📦',layout='wide',initial_sidebar_state='collapsed')
CSS="""
<style>
:root{--gc-navy:#153c4a;--gc-teal:#087f8c;--gc-muted:#607b86}
.block-container{max-width:1080px;padding-top:1.15rem;padding-bottom:3rem}
[data-testid="stAppViewContainer"]{background:#f5f9fa}
[data-testid="stHeader"]{background:rgba(245,249,250,.94)}
.gc-brand{display:flex;align-items:center;gap:11px;margin:0 0 16px}
.gc-logo{font-weight:800;font-size:27px;letter-spacing:.4px;color:#153c4a}.gc-logo span{color:#087f8c}
.gc-tagline{font-size:12px;color:#607b86;border-left:1px solid #c7dade;padding-left:11px;line-height:1.5}
.gc-title{font-size:27px;line-height:1.45;font-weight:750;color:#153c4a;margin:4px 0 7px}
.gc-lead{font-size:15px;line-height:1.8;color:#526e79;margin-bottom:16px}
.gc-step{display:flex;flex-wrap:wrap;gap:8px;margin:0 0 16px}.gc-step span{font-size:12px;background:#eaf1f3;color:#607b86;padding:6px 10px;border-radius:20px}.gc-step .done{background:#dcf0ea;color:#216f5d}.gc-step .current{background:#dff0f3;color:#086875;font-weight:650}
.gc-summary{display:grid;grid-template-columns:1fr 1fr;gap:11px;margin:5px 0 17px}
.gc-summary>div{background:#fff;border:1px solid #dce8eb;padding:16px;border-radius:14px}
.gc-summary .gc-value{display:block;font-size:30px;font-weight:750;line-height:1.3;color:#153c4a;margin-top:4px}
.gc-summary .gc-caption{font-size:12px;color:#607b86}.gc-summary small{font-size:11px;color:#607b86}
.gc-card{background:#fff;border:1px solid #dbe7e9;border-radius:14px;padding:18px;margin:4px 0 5px}
.gc-card.red{border-left:5px solid #cb4c52}.gc-card.amber{border-left:5px solid #d5a243}.gc-card.green{border-left:5px solid #3a9d83}.gc-card.gray{border-left:5px solid #98a8ae}
.gc-card-top{display:flex;justify-content:space-between;gap:8px;align-items:start}.gc-name{font-size:19px;line-height:1.5;font-weight:700;color:#153c4a}.gc-id{font-size:11px;color:#78909a}
.gc-badge{font-size:11px;display:inline-block;border-radius:15px;padding:5px 9px;white-space:nowrap}.gc-badge.red{background:#fce8e8;color:#aa3c47}.gc-badge.amber{background:#fff2d8;color:#91621c}.gc-badge.green{background:#e1f3ec;color:#246e59}.gc-badge.gray{background:#edf0f2;color:#687a83}
.gc-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:16px}.gc-field{font-size:12px;color:#607b86}.gc-field strong{display:block;color:#153c4a;font-size:19px;margin-top:3px;line-height:1.6}.gc-field em{font-size:12px;font-style:normal;color:#78909a}
.gc-date{background:#eff7f8;border-radius:9px;padding:10px 12px;margin-top:13px;font-size:13px;line-height:1.8;color:#234e5a}.gc-date.late{background:#fff0ee;color:#ad3d49}
.gc-small{font-size:12px;line-height:1.8;color:#617d87}.gc-note{background:#eaf4f6;color:#305f6b;padding:12px 15px;border-radius:10px;font-size:13px;line-height:1.85;margin:8px 0 14px}
.gc-section{font-size:20px;font-weight:700;color:#153c4a;margin:18px 0 6px}.gc-foot{border-top:1px solid #d8e5e8;margin-top:24px;padding-top:12px;font-size:11px;color:#77909b;line-height:1.9}
.gc-empty{text-align:center;padding:25px;background:#fff;border:1px dashed #c1d7dd;border-radius:12px;color:#526f78}
[data-testid="stFileUploader"]{background:white;border:1px solid #d6e5e8;border-radius:12px;padding:12px}
[data-testid="stRadio"] div[role="radiogroup"]{gap:5px;flex-wrap:wrap}
[data-testid="stRadio"] label{background:#fff;border:1px solid #d7e4e8;border-radius:9px;padding:7px 9px!important;margin:0!important}
[data-testid="stExpander"]{background:#fff;border-radius:11px}
h1,h2,h3{color:#153c4a}
@media(max-width:600px){.block-container{padding:1rem 1rem 3rem}.gc-title{font-size:23px}.gc-logo{font-size:25px}.gc-tagline{font-size:10px}.gc-card{padding:14px}.gc-name{font-size:17px}.gc-summary>div{padding:13px}.gc-summary .gc-value{font-size:27px}.gc-field strong{font-size:18px}.gc-badge{font-size:10px;padding:4px 7px}}
</style>
"""
st.markdown(CSS,unsafe_allow_html=True)
def html(x):st.markdown(x,unsafe_allow_html=True)
def txt(x):return escape(str(x if x is not None else ''))
def valid(x):
 try:return x is not None and not pd.isna(x) and math.isfinite(float(x))
 except (TypeError,ValueError):return False
def num(x,dec=1):
 if not valid(x):return 'ยังไม่มีข้อมูล'
 return f'{float(x):,.{dec}f}'.rstrip('0').rstrip('.') if dec else f'{float(x):,.0f}'
def showtime(x):
 if x is None or str(x) in ('','None','nan','NaT'):return 'ยังไม่พบในช่วงแผน'
 try:return stamp(x).strftime('%d/%m/%Y %H:%M')+' น.'
 except Exception:return str(x)
def thai_status(code):return {'SHORT':'ขาดตามแผน','LOW':'ต่ำกว่าสำรอง','REVIEW':'ควรทบทวนเติมเต็ม','OK':'ไม่พบขาดในช่วงนี้','DATA':'ข้อมูลยังไม่ครบ'}.get(code,code)
def color(code):return {'SHORT':'red','LOW':'amber','REVIEW':'amber','OK':'green','DATA':'gray'}.get(code,'gray')
def section(title,description=''):
 html('<div class="gc-section">'+txt(title)+'</div>')
 if description:st.caption(description)
def note(message):html('<div class="gc-note">'+txt(message)+'</div>')
def summary(items):
 html('<div class="gc-summary">'+''.join('<div><span class="gc-caption">'+txt(label)+'</span><span class="gc-value">'+txt(value)+'</span><small>'+txt(caption)+'</small></div>' for label,value,caption in items)+'</div>')
def safe_frame(rows):
 df=pd.DataFrame(rows)
 for col in df.columns:df[col]=df[col].map(lambda x:'; '.join(str(i) for i in x) if isinstance(x,list) else x)
 return df
def thai_frame(rows,cols):
 df=safe_frame(rows)
 return df.reindex(columns=list(cols)).rename(columns=cols)
def report_bytes(result):
 forecast=[]
 for f in result.get('forecasts',[]):
  for m in f.get('methods',[]):forecast.append({'sku':f['sku'],'target_month':f.get('target'),'selected':m.get('method')==f.get('selected',{}).get('method'),**{k:v for k,v in m.items() if k!='errors'}})
 tables={'Forecast':safe_frame(forecast),'Materials':safe_frame(result.get('materials',[])),'Orders':safe_frame(result.get('orders',[])),'Requirements':safe_frame(result.get('requirements',[])),'Timeline':safe_frame(result.get('timeline',[])),'Delivery_KPI':pd.DataFrame([result.get('kpi',{})])}
 buf=BytesIO()
 with pd.ExcelWriter(buf,engine='openpyxl') as writer:
  for name,df in tables.items():
   frame=df.copy()
   for col in frame.columns:frame[col]=frame[col].map(lambda x:"'"+x if isinstance(x,str) and x.lstrip().startswith(('=','+','-','@')) else x)
   frame.to_excel(writer,sheet_name=name,index=False)
 return buf.getvalue()

html('<div class="gc-brand"><div class="gc-logo">Gate<span>CAT</span></div><div class="gc-tagline">วัสดุพร้อม ผลิตได้<br>มองความเสี่ยงก่อนถึงรอบส่ง</div></div>')
html('<div class="gc-title">แผนนี้ต้องเตรียมวัสดุอะไร?</div><div class="gc-lead">เลือก Excel เพียงไฟล์เดียว เว็บจะช่วยคาดความต้องการและหาวัสดุที่เสี่ยงไม่พอ</div>')
upload=st.file_uploader('① เลือก Excel ข้อมูลเข้า (.xlsx) ไม่เกิน 5MB',type=['xlsx'],key='gatecat_excel_upload')
base=Path(__file__).resolve().parent
sample_path=base/'GateCAT_simulated_inputs.xlsx'
blank_path=base/'GateCAT_blank_template.xlsx'
if upload is not None:
 if upload.size>5*1024*1024:st.error('ไฟล์ใหญ่เกิน5MBสำหรับต้นแบบ กรุณาใช้ไฟล์ตามแบบฟอร์ม');st.stop()
 raw=upload.getvalue();source=BytesIO(raw);fingerprint=hashlib.sha256(raw).hexdigest()[:12];source_label=upload.name
elif sample_path.exists():
 raw=sample_path.read_bytes();source=BytesIO(raw);fingerprint='demo';source_label='ข้อมูลตัวอย่างที่มากับเว็บ'
else:
 html('<div class="gc-step"><span class="current">1 เลือก Excel</span><span>2 ดูวัสดุที่ต้องเตรียม</span><span>3 ตรวจรอบส่ง</span></div>')
 note('เริ่มจาก Excel ข้อมูลจำลองที่ชื่อขึ้นต้น GateCAT_simulated_inputs ไม่เลือกแบบฟอร์มเปล่าหรือ reference_results')
 section('เว็บจะตอบให้คุณ 4 เรื่อง')
 st.write('• วัสดุอะไรมีความเสี่ยงไม่พอ\n• ต้องเตรียมเพิ่มตามแผนเท่าไร\n• ควรสั่งช้าที่สุดวันและเวลาไหน\n• ออร์เดอร์ใดอาจได้รับผลกระทบ')
 st.caption('ยังไม่ต้องกรอกสูตรหรือแก้โค้ด เลือกไฟล์ด้านบนเพื่อเริ่ม')
 st.stop()
try:
 with st.spinner('กำลังอ่าน Excel และตรวจข้อมูล…'):data,input_tabs=read_workbook(source)
except Exception as e:
 st.error('ยังอ่านไฟล์นี้ไม่ได้: '+str(e));note('ใช้ Excel ข้อมูลจำลอง/แบบฟอร์ม GateCAT และรักษาชื่อชีตกับหัวคอลัมน์เดิม ข้อความด้านบนระบุจุดที่ต้องตรวจ');st.stop()
with st.expander('ไฟล์ที่ใช้และการตั้งค่าแผน',expanded=False):
 st.caption('ไฟล์: '+str(source_label))
 st.write('เวลาตัดข้อมูล: '+showtime(data['Settings']['as_of'])+' (+07:00)')
 st.caption('วันที่แสดงผลอ้างอิง snapshot ใน Excel ไม่ใช่วันที่ปัจจุบันโดยอัตโนมัติ')
 horizon=st.number_input('วางแผนล่วงหน้ากี่วัน',min_value=1,max_value=180,value=int(data['Settings']['horizon_days']),step=1,key='h_'+fingerprint)
 include=st.checkbox('เตรียม Forecast ส่วนที่ยังไม่เป็น Order ด้วย',value=bool(data['Settings']['include_forecast']),key='f_'+fingerprint)
 st.caption('รุ่นนี้คาด Demand หนึ่งเดือนเป้าหมาย และรวม Orders ในขอบเขตที่มีข้อมูล')
data=copy.deepcopy(data);data['Settings']['horizon_days']=int(horizon);data['Settings']['include_forecast']=bool(include)
try:
 with st.spinner('กำลังพยากรณ์และตรวจความพร้อมวัสดุ…'):result=calculate(data)
except Exception as e:st.error('ยังคำนวณไม่ได้: '+str(e));st.stop()
label=str(data['Settings'].get('data_label',''))
st.caption(label)
html('<div class="gc-step"><span class="done">✓ อ่าน Excel แล้ว</span><span class="current">2 ดูสิ่งที่ต้องเตรียม</span><span>3 ตรวจรอบส่ง</span></div>')
materials=list(result.get('materials',[]));orders=list(result.get('orders',[]))
priority={'SHORT':0,'LOW':1,'REVIEW':2,'DATA':3,'OK':4}
sorted_materials=sorted(materials,key=lambda m:(priority.get(m.get('status'),9),str(m.get('latest_order_at') or '9999'),m['material_id']))
def affected_orders(mid):
 return [o for o in orders if any(str(r).split(' ')[0].rstrip(':')==mid for r in o.get('material_risks',[]))]
def material_card(m):
 code=m.get('status','DATA');cl=color(code);unit=str(m.get('unit',''));buy=m.get('suggest_order_qty',0);late=bool(m.get('normal_order_late'));short=m.get('first_shortage');latest=m.get('latest_order_at');name=str(m.get('name',m['material_id']));need_action=code in ('SHORT','LOW','REVIEW')
 buylabel='ปริมาณเสนอให้ทบทวนจัดหา' if valid(buy) and buy>0 else 'ปริมาณเสนอซื้อในช่วงนี้'
 buyvalue=num(buy)+' '+unit if valid(buy) else 'ยังประเมินไม่ได้'
 buydate=showtime(latest) if latest else 'ยังไม่มีวันเสนอในช่วงนี้'
 if late:buydate+=' · วันสั่งปกติผ่านแล้ว'
 html('<div class="gc-card '+cl+'"><div class="gc-card-top"><div><div class="gc-name">'+txt(name)+'</div><div class="gc-id">'+txt(m['material_id'])+'</div></div><span class="gc-badge '+cl+'">'+txt(thai_status(code))+'</span></div><div class="gc-grid"><div class="gc-field">คงเหลืออิสระตั้งต้น<strong>'+txt(num(m.get('free_stock')))+' '+txt(unit)+'</strong></div><div class="gc-field">'+txt(buylabel)+'<strong>'+txt(buyvalue)+'</strong></div></div><div class="gc-date '+('late' if late else '')+'">คาดเริ่มขาด: '+txt(showtime(short))+'<br>ควรสั่งช้าที่สุด: '+txt(buydate)+'</div></div>')
 if late:st.caption('วันสั่งปกติผ่านเมื่อเทียบกับ snapshot ต้องทบทวนการจัดหา/แผน ไม่สั่งย้อนหลังและสมมติว่าของมาทัน')
 impacted=affected_orders(m['material_id'])
 with st.expander('ดูที่มาและออร์เดอร์ที่เกี่ยวข้อง · '+m['material_id']):
  st.write('วัสดุรวมที่แผนต้องใช้: '+num(m.get('gross_requirement'))+' '+unit)
  st.write('ระดับสำรอง (SS): '+num(m.get('SS'))+' '+unit+' · จุดทบทวนสั่งซื้อ (ROP): '+num(m.get('ROP'))+' '+unit)
  st.write('ระยะเวลาจัดหา: '+num(m.get('lead_days'))+' วันปฏิทิน')
  st.caption('ปริมาณเสนอครอบคลุมช่วงแผน '+str(horizon)+' วันและเติมถึง SS ปัดตาม pack/MOQ ไม่ใช่ปริมาณขาดของวันเดียวและไม่ใช่คำสั่งซื้อจริง')
  if impacted:
   st.write('ออร์เดอร์ที่มีข้อเตือนวัสดุรายการนี้')
   st.dataframe(thai_frame(impacted,{'order_id':'ออร์เดอร์','customer':'ลูกค้า','due_at':'วันเวลาที่ต้องถึง','backlog':'ยังไม่ส่ง (ชิ้น)','good_to_make':'สินค้าดีต้องผลิตเพิ่ม'}),hide_index=True,use_container_width=True)
  else:st.caption('ไม่มีออร์เดอร์ตรงที่พบข้อเตือนขาด อาจเป็นการเตรียม Forecast/ระดับสำรอง หรือข้อมูลยังไม่ครบ')
  lines=[x for x in result.get('timeline',[]) if x.get('material_id')==m['material_id']]
  st.write('เส้นเวลารับ–ใช้วัสดุ')
  st.dataframe(thai_frame(lines,{'at':'วันเวลา','kind':'ประเภท','reference':'อ้างอิง','receipt_qty':'คาดรับ','need_qty':'ต้องใช้','projected_balance':'คงเหลือคาด'}),hide_index=True,use_container_width=True)
  st.caption('RECEIPT = คาดรับ · FIRM = งาน Order · BUFFER = แผนสำรอง · SNAPSHOT = ยอดตั้งต้น; ของคาดรับยังไม่ใช่ของรับจริง')

def go(page_name):st.session_state['gatecat_page']=page_name
PAGES=['เริ่มดูผล','วัสดุ / สั่งซื้อ','พยากรณ์สินค้า','Orders / FG','ข้อมูล / รายงาน']
if st.session_state.get('gatecat_page') not in PAGES:st.session_state['gatecat_page']=PAGES[0]
page=st.radio('เลือกเรื่องที่ต้องการดู',PAGES,horizontal=True,key='gatecat_page')
if page=='เริ่มดูผล':
 section('เริ่มจากวัสดุที่ต้องเตรียม','ยังไม่ต้องอ่านทุกตาราง ดูรายการเสี่ยงก่อนหนึ่งรายการ')
 short_count=sum(m.get('status')=='SHORT' for m in materials);review_count=sum(m.get('status') in ('LOW','REVIEW') for m in materials)
 summary([('วัสดุเสี่ยงขาด',str(short_count)+' รายการ','ต้องตรวจวัสดุพร้อมตามเวลา'),('ควรทบทวนเติมเต็ม',str(review_count)+' รายการ','ต่ำกว่าสำรองหรือถึง ROP')])
 attention=[m for m in sorted_materials if m.get('status') in ('SHORT','LOW','REVIEW','DATA')]
 if attention:
  note('อ่านการ์ดจากซ้ายไปขวา: ของที่มี → จำนวนเสนอเตรียมเพิ่ม → วันควรสั่ง ค่าทั้งหมดอ้างอิงไฟล์และช่วงแผนที่เลือก')
  for m in attention[:2]:material_card(m)
  st.button('ดูวัสดุทั้งหมด ('+str(len(materials))+')',on_click=go,args=('วัสดุ / สั่งซื้อ',),type='primary',use_container_width=True)
 else:note('ไม่พบวัสดุขาดในข้อมูลและช่วงแผนนี้ ยังไม่ใช่การรับประกันว่าทุกงานส่งทัน')
 with st.expander('อยากดูว่าต้องทำอะไรต่อ?'):
  st.write('1. ตรวจรายการวัสดุที่ต้องเตรียม\n2. เปิด “ดูที่มา” ถ้าต้องการรู้ว่ากระทบออร์เดอร์ใด\n3. ใช้ผลเป็นข้อมูลทบทวนกับผู้วางแผน/จัดซื้อ\n4. แก้ Excel เมื่อข้อมูลเปลี่ยน แล้วอัปโหลดใหม่ เว็บคำนวณใหม่ให้')
  st.caption('เว็บรุ่นนี้ไม่สั่งซื้อ ไม่เพิ่มสต็อก และไม่เปลี่ยนผลรับมอบจริงเอง')
elif page=='วัสดุ / สั่งซื้อ':
 section('วัสดุอะไรควรเตรียมเพิ่ม?','เลือกหนึ่งรายการแล้วอ่านจำนวนและวันบนการ์ด ไม่ต้องเลื่อนตารางกว้าง')
 choice=st.selectbox('เลือกวัสดุ',['ทั้งหมดที่ควรดู']+[m['material_id']+' · '+str(m.get('name','')) for m in sorted_materials],key='mat_'+fingerprint)
 shown=sorted_materials if choice=='ทั้งหมดที่ควรดู' else [m for m in sorted_materials if choice.startswith(m['material_id']+' · ')]
 for m in shown:material_card(m)
 with st.expander('ความหมายของผลที่แสดง'):
  st.write('ขาดตามแผน: คงเหลือคาดติดลบเมื่อถึงเวลาต้องใช้\nต่ำกว่าสำรอง: ยังไม่จำเป็นต้องติดลบ แต่ต่ำกว่ากันชน\nทบทวนเติมเต็ม: เข้าเงื่อนไข ROP\nไม่พบขาด: ไม่พบจากข้อมูล/ช่วงที่กำหนด ไม่รับประกันผลหน้างาน')
  st.caption('สีเป็นกฎการเตือน ไม่ใช่เปอร์เซ็นต์ความน่าจะเป็น')
elif page=='พยากรณ์สินค้า':
 section('เว็บคาดว่าต้องเตรียมสินค้าดีเท่าไร?','นี่คือความต้องการสินค้า ไม่ใช่จำนวนวัตถุดิบที่ต้องสั่งตรง ๆ')
 forecasts=result.get('forecasts',[])
 if not forecasts:note('ยังไม่มีข้อมูลพยากรณ์')
 else:
  fsku=st.selectbox('เลือกสินค้า',[x['sku'] for x in forecasts],key='forecast_'+fingerprint)
  f=next(x for x in forecasts if x['sku']==fsku)
  if f.get('error'):st.warning(f['error'])
  elif not f.get('selected'):st.warning('ข้อมูลผลพยากรณ์ยังไม่ครบ')
  else:
   selected=f['selected'];summary([('คาดความต้องการเดือน '+str(f.get('target',result.get('target_month',''))),num(selected.get('forecast'))+' ชิ้นดี','คำนวณจากประวัติใน Excel'),('วิธีที่เว็บเลือก',selected.get('method','—'),'เลือกตาม MAD ต่ำสุด')])
   history=[{'เดือน':x['period'],'ความต้องการจริง':x['demand_qty']} for x in f.get('history',[])]
   if history:st.line_chart(pd.DataFrame(history).set_index('เดือน'))
   note('ขั้นถัดไปเว็บจะหัก FG และแปลงผ่าน BOM ไปเป็นวัสดุให้เอง คุณไม่ต้องนำเลขนี้ไปคูณสูตรด้วยมือ')
   with st.expander('ดูสามวิธีและสูตรที่ใช้ (รายละเอียดงานวิจัย)'):
    st.dataframe(thai_frame(f.get('methods',[]),{'method':'วิธี','forecast':'Forecast เป้าหมาย','MAD':'MAD','MSE':'MSE','RMSE':'RMSE','MAPE':'MAPE (%)','n':'จุดประเมิน','n_valid':'จุด MAPE'}),hide_index=True,use_container_width=True)
    st.code('SMA3 = (A[t-1]+A[t-2]+A[t-3])/3\nWMA3 = 0.50*A[t-1]+0.30*A[t-2]+0.20*A[t-3]\nSES = 0.30*A[t-1]+0.70*F[t-1], F2=A1',language='text')
    st.caption('ช่วงประเมินเดียวกัน MAPE ข้าม Actual=0; ค่าพยากรณ์ไม่ใช่การรับประกันยอดลูกค้าจริง')
elif page=='Orders / FG':
 section('ออร์เดอร์ไหนยังต้องเตรียมของ?','จำนวนยังไม่ส่งรวมทั้งออร์เดอร์ที่ยังไม่ถึงกำหนด ไม่ได้แปลว่าช้าทั้งหมด')
 if not orders:note('ไม่มี Orders ในไฟล์นี้')
 else:
  oid=st.selectbox('เลือกออร์เดอร์',[o['order_id']+' · '+str(o.get('customer','')) for o in orders],key='ord_'+fingerprint)
  o=next(o for o in orders if oid.startswith(o['order_id']+' · '))
  summary([('ยังไม่ได้ส่ง',num(o.get('backlog'),0)+' ชิ้น','จากยอด '+num(o.get('qty'),0)+' ชิ้น'),('FG จัดสรรในแผน',num(o.get('fg_allocated_plan'),0)+' ชิ้น','ยังไม่ใช่การส่งจริง')])
  st.write('ต้องถึงลูกค้าเดิม: '+showtime(o.get('due_at')))
  st.write('สินค้าดีที่แผนต้องผลิตเพิ่ม: '+num(o.get('good_to_make'),0)+' ชิ้น')
  if o.get('material_risks'):
   st.warning('มีวัสดุที่ควรตรวจ')
   for risk in o['material_risks']:st.write('• '+str(risk))
  elif o.get('readiness')=='FG_READY':st.success('มี FG รองรับตามแผน แต่ยังต้องยืนยันการส่งจริง')
  elif o.get('readiness')=='CLOSED':st.info('รายการนี้ส่งครบตามประวัติแล้ว')
  else:note('ยังไม่พบขาดในวัสดุที่คำนวณได้ แต่รุ่นนี้ยังไม่ยืนยันคิวเครื่อง/วันส่งจริง')
  with st.expander('ดูผลส่งจริงและภาพรวมการส่งมอบ'):
   st.write('รับจริงแล้ว '+num(o.get('delivered'),0)+' ชิ้น · ถึงทันเวลา '+num(o.get('on_time_qty'),0)+' ชิ้น')
   st.write('เวลาที่ส่งครบจริง: '+(showtime(o.get('full_at')) if o.get('full_at') else 'ยังส่งไม่ครบ'))
   st.write('ช้าหลังส่งครบ: '+(num(o.get('late_minutes'))+' นาที' if valid(o.get('late_minutes')) else 'ยังไม่มีเวลาส่งครบให้คำนวณ'))
   k=result.get('kpi',{});summary([('รอบส่งครบตรงเวลา',(num(k.get('OTD_calloff'))+'%' if valid(k.get('OTD_calloff')) else 'ยังไม่มีฐานประเมิน'),'เฉพาะ '+str(k.get('due_count',0))+' รอบที่ถึงกำหนด'),('ยอดครบสะสม',(num(k.get('cumulative_fulfilment'))+'%' if valid(k.get('cumulative_fulfilment')) else 'ยังไม่มีฐานประเมิน'),'รวมส่งชดเชย ณ snapshot')])
   st.caption('ส่งชดเชยครบ ไม่ล้างประวัติว่ารอบเดิมส่งช้า')
 with st.expander('ดู FG คงเหลือและรายการ Orders ทั้งหมด'):
  st.dataframe(thai_frame(data.get('FG_Stock',[]),{'sku':'สินค้า','on_hand':'FG กายภาพ','reserved_external':'จองภายนอก','safety_target':'FG เป้าหมายขั้นต่ำ'}),hide_index=True,use_container_width=True)
  st.dataframe(thai_frame(orders,{'order_id':'ออร์เดอร์','sku':'สินค้า','qty':'ต้องส่ง','due_at':'กำหนดเดิม','delivered':'รับแล้ว','backlog':'ยังไม่ส่ง','good_to_make':'ต้องผลิตดีเพิ่ม','readiness':'สถานะแผน'}),hide_index=True,use_container_width=True)
elif page=='ข้อมูล / รายงาน':
 section('นำผลไปใช้ต่อ')
 st.download_button('ดาวน์โหลดผลคำนวณเป็น Excel',report_bytes(result),file_name='GateCAT_forecast_and_material_plan.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary',use_container_width=True)
 if sample_path.exists():st.download_button('ดาวน์โหลด Excel ข้อมูลจำลอง',sample_path.read_bytes(),file_name=sample_path.name)
 if blank_path.exists():st.download_button('ดาวน์โหลดแบบฟอร์มเปล่า',blank_path.read_bytes(),file_name=blank_path.name)
 with st.expander('ดูข้อมูลที่อ่านจาก Excel'):
  sheet=st.selectbox('เลือกชีต',list(input_tabs.keys()),key='sheet_'+fingerprint);st.dataframe(input_tabs[sheet],hide_index=True,use_container_width=True)
 with st.expander('ดูที่มาของจำนวนวัสดุ'):
  st.dataframe(thai_frame(result.get('requirements',[]),{'job_id':'งาน','sku':'สินค้า','material_id':'วัสดุ','good':'สินค้าดีต้องผลิต','start_qty':'จำนวนเริ่มผลิต','bom_per_start':'BOM ต่อชิ้นเริ่ม','run_material':'วัสดุ Run','startup_loss':'วัสดุเริ่มงาน','gross_requirement':'รวมวัสดุ','use_at':'เวลาต้องใช้'}),hide_index=True,use_container_width=True)
 with st.expander('ข้อจำกัดและสมมติฐาน'):
  st.write('• อ่าน Excel snapshot แล้ววางแผน ไม่ใช่ฐานข้อมูลหลายผู้ใช้หรือคลังปรับยอดอัตโนมัติ\n• วันต้องใช้มาจาก Due ลบ lead time ที่กรอก ยังไม่จัดคิวเครื่อง/Changeover จริง\n• SS สูตรง่ายใช้ Lead time คงที่และการใช้รายวันอิสระ\n• หนึ่ง Order ที่ผลิตใหม่คิดเป็นหนึ่งล็อต Startup\n• มีคำแนะนำสั่งไม่ได้แปลว่าได้รับของแล้ว\n• ไม่รับประกันส่งทันหรือคุณภาพของบริษัทจริง')
if result.get('warnings'):
 with st.expander('ข้อมูลที่ควรตรวจเพิ่มเติม ('+str(len(result['warnings']))+')'):
  for w in result['warnings']:st.warning(w)
html('<div class="gc-foot">GateCAT · UI สำหรับดูคำตอบทีละรายการบนโทรศัพท์<br>คำนวณด้วย engine.py ชุดเดิม · ไม่มีการสั่งซื้อจริงอัตโนมัติ · ใช้ข้อมูลที่ได้รับอนุญาตเท่านั้น</div>')
