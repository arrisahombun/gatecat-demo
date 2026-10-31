
from pathlib import Path
from io import BytesIO
import hashlib,copy
import pandas as pd
import streamlit as st
from engine import read_workbook,calculate,stamp
st.set_page_config(page_title='GateCAT | Excel Forecast',page_icon='📦',layout='wide')
st.markdown("<style>.block-container{padding-top:2rem} h1,h2,h3{color:#12384c}</style>",unsafe_allow_html=True)
st.title('GateCAT · อัปโหลด Excel แล้วพยากรณ์และวางแผนวัสดุ')
st.caption('Demand → FG → จำนวนเริ่มผลิต → BOM/Material → วันควรสั่ง → รอบส่งที่มีความเสี่ยง')
base=Path(__file__).resolve().parent
with st.sidebar:
 st.header('ข้อมูลเข้า')
 upload=st.file_uploader('อัปโหลดสมุดงาน GateCAT (.xlsx)',type=['xlsx'])
 sample=base/'GateCAT_simulated_inputs.xlsx';template=base/'GateCAT_blank_template.xlsx'
 if sample.exists():st.download_button('ดาวน์โหลด Excel ตัวอย่าง',sample.read_bytes(),file_name=sample.name,mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
 if template.exists():st.download_button('ดาวน์โหลดแบบฟอร์มเปล่า',template.read_bytes(),file_name=template.name,mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
 st.info('ใช้ไฟล์เดียว 11 ชีตข้อมูลหลัก อ่านวิธีกรอกใน START_HERE และ Data_Dictionary ไม่ต้องคำนวณ Forecast เองใน Excel')
if upload:
 if upload.size>5*1024*1024:st.error('ไฟล์เกิน5MB ซึ่งเกินขอบเขตต้นแบบ');st.stop()
 source=BytesIO(upload.getvalue());source_name=upload.name;fingerprint=hashlib.sha256(upload.getvalue()).hexdigest()[:12]
elif sample.exists():source=BytesIO(sample.read_bytes());source_name='ข้อมูลจำลองตัวอย่าง';fingerprint='demo'
else:st.info('กรุณาอัปโหลด Excel เพื่อเริ่มคำนวณ');st.stop()
try:data,input_tabs=read_workbook(source)
except Exception as e:st.error('อ่าน/ตรวจ Excel ไม่สำเร็จ: '+str(e));st.stop()
with st.sidebar:
 st.subheader('ขอบเขตแผน')
 horizon=st.number_input('จำนวนวันของแผน',min_value=1,max_value=180,value=int(data['Settings']['horizon_days']),step=1,key='h_'+fingerprint)
 include=st.checkbox('รวม Forecast ส่วนไม่ทับ Order',value=data['Settings']['include_forecast'],key='f_'+fingerprint)
 st.caption('เวลาตัดข้อมูล '+data['Settings']['as_of']+' · แก้ใน Settings หากเปลี่ยน snapshot')
 data=copy.deepcopy(data);data['Settings']['horizon_days']=horizon;data['Settings']['include_forecast']=include
try:result=calculate(data)
except Exception as e:st.error('คำนวณไม่สำเร็จ: '+str(e));st.stop()
st.warning(str(data['Settings']['data_label']))
st.caption('ไฟล์ '+source_name+' · Forecast เดือน '+result['target_month']+' · เวลาทั้งหมดตีความ +07:00 · คำนวณอัตโนมัติเมื่ออัปโหลด/เปลี่ยนขอบเขตแผน')
materials=pd.DataFrame(result['materials']);orders=pd.DataFrame(result['orders']);requirements=pd.DataFrame(result['requirements']);timeline=pd.DataFrame(result['timeline'])
forecast_rows=[]
for f in result['forecasts']:
 for m in f['methods']:forecast_rows.append({'sku':f['sku'],'target_month':f['target'],'selected':m['method']==f['selected']['method'],**{k:v for k,v in m.items() if k!='errors'}})
forecasts=pd.DataFrame(forecast_rows)
def show_number(x):return '—' if x is None or pd.isna(x) else f'{x:,.2f}'
def safe_report(frames):
 out=BytesIO()
 with pd.ExcelWriter(out,engine='openpyxl') as writer:
  for name,df in frames.items():
   frame=df.copy()
   for col in frame.columns:
    frame[col]=frame[col].map(lambda x: "; ".join(map(str,x)) if isinstance(x,list) else x)
    frame[col]=frame[col].map(lambda x: "'"+x if isinstance(x,str) and x.lstrip().startswith(('=','+','-','@')) else x)
   frame.to_excel(writer,sheet_name=name,index=False)
 return out.getvalue()
LABEL={'SHORT':'ขาดตามแผน','LOW':'ต่ำกว่าสำรอง','REVIEW':'ทบทวนตาม ROP','OK':'ไม่พบขาดในขอบเขต','DATA':'ข้อมูลคำนวณ SS/ROP ไม่ครบ'}
tabs=st.tabs(['ภาพรวม','พยากรณ์อัตโนมัติ','วัสดุ/วันสั่งซื้อ','FG/Orders','เส้นเวลาและสูตร','ข้อมูลเข้า/ส่งออก'])
with tabs[0]:
 c=st.columns(4);c[0].metric('วัสดุขาดตามแผน',int((materials['status']=='SHORT').sum()));c[1].metric('ต่ำกว่าสำรอง',int((materials['status']=='LOW').sum()));c[2].metric('Order วัสดุเสี่ยง',sum(x['readiness']=='MATERIAL_RISK' for x in result['orders']));c[3].metric('ค้างส่งรวม (ชิ้น)',f"{result['kpi']['backlog_all']:,.0f}")
 st.subheader('วัสดุที่ควรดำเนินการ')
 display=materials[['material_id','name','unit','free_stock','SS','ROP','first_shortage','latest_order_at','suggest_order_qty','normal_order_late','status']].copy();display['status']=display['status'].map(LABEL)
 st.dataframe(display,use_container_width=True,hide_index=True)
 c=st.columns(3);c[0].metric('รอบส่งครบตรงเวลา',show_number(result['kpi']['OTD_calloff'])+'%');c[1].metric('จำนวนที่ถึงทันเวลา',show_number(result['kpi']['on_time_quantity'])+'%');c[2].metric('ยอดครบสะสมภายหลัง',show_number(result['kpi']['cumulative_fulfilment'])+'%')
 st.caption('KPI นี้ใช้เฉพาะ Orders ที่ถึงกำหนดแล้ว ณ as_of จำนวน '+str(result['kpi']['due_count'])+' รายการ; ส่งชดเชยครบไม่ย้อนล้างความล่าช้าเดิม')
 if result['warnings']:
  with st.expander('ข้อมูลที่ต้องตรวจ ('+str(len(result['warnings']))+')'):
   for w in result['warnings']:st.warning(w)
with tabs[1]:
 st.subheader('เว็บคำนวณสามวิธีและเลือกตาม MAD ให้เอง')
 sku=st.selectbox('เลือกสินค้า',[p['sku'] for p in data['Products']],key='forecast_sku')
 f=next(x for x in result['forecasts'] if x['sku']==sku)
 if f.get('error'):st.error(f['error'])
 else:
  c=st.columns(3);c[0].metric('วิธีที่เลือก',f['selected']['method']);c[1].metric('Demand เดือนเป้าหมาย (ชิ้นดี)',show_number(f['selected']['forecast']));c[2].metric('MAD',show_number(f['selected']['MAD']))
  st.dataframe(forecasts[forecasts['sku']==sku],use_container_width=True,hide_index=True)
  chart=pd.DataFrame([{'period':x['period'],'Demand จริง':x['demand_qty'],'Forecast เป้าหมาย':None} for x in f['history']]+[{'period':f['target'],'Demand จริง':None,'Forecast เป้าหมาย':f['selected']['forecast']}]).set_index('period');st.line_chart(chart)
  st.write('Orders ทั้งเดือน:',f['firm_month'],'· Forecast ส่วนไม่ทับ Order:',round(f['forecast_residual'],2),'· Good ใหม่สำหรับบัฟเฟอร์:',f['buffer_good'])
  with st.expander('ค่าพยากรณ์ย้อนหลังสำหรับตรวจสูตร'):
   for m in f['methods']:st.write(m['method']);st.dataframe(pd.DataFrame(m['errors']),hide_index=True,use_container_width=True)
 st.info('เลือก MAD ต่ำสุด ตามด้วย MSE และลำดับ SMA/WMA/SES ถ้าเสมอ · SMA3 / WMA0.50,0.30,0.20 / SES α0.30 · MAPE เมื่อ Actual0 เป็น N/A')
with tabs[2]:
 st.subheader('จะขาดเมื่อไร ควรสั่งช้าที่สุดวันไหน')
 st.dataframe(display,use_container_width=True,hide_index=True)
 mid=st.selectbox('ดูรายละเอียดวัสดุ',materials['material_id'].tolist(),key='material_detail')
 m=next(x for x in result['materials'] if x['material_id']==mid)
 st.write('สถานะ:',LABEL[m['status']],'· คงเหลืออิสระ:',m['free_stock'],m['unit'])
 st.write('ต่ำกว่า SS ครั้งแรก:',m['first_below_SS'],'· ถึง0:',m['first_zero'],'· ขาด:',m['first_shortage'])
 st.write('วันควรสั่งช้าที่สุด:',m['latest_order_at'],'· ปริมาณเสนอ:',m['suggest_order_qty'],m['unit'])
 if m['normal_order_late']:st.error('ผ่านวันสั่งตาม Lead time ปกติแล้ว ต้องทบทวนการจัดหา/รอบส่ง ไม่สมมติว่าสั่งย้อนหลังได้')
 st.dataframe(timeline[timeline['material_id']==mid],hide_index=True,use_container_width=True)
 st.caption('ใบรับเป็นคาดการณ์ ไม่ใช่ของรับจริง; ปริมาณเสนอครอบคลุมช่วงแผนและเติมถึง SS ปัดตาม pack/MOQ ไม่ใช่ EOQ และยังไม่สร้าง PO อัตโนมัติ')
with tabs[3]:
 st.subheader('สินค้าสำเร็จรูปและการจัดสรรให้ Orders')
 fgt=pd.DataFrame(data['FG_Stock']);fgt['free_start']=fgt['on_hand']-fgt['reserved_external'];fgt['unallocated_after_plan']=fgt['sku'].map(result['fg_remaining_plan']);st.dataframe(fgt,hide_index=True,use_container_width=True)
 st.dataframe(orders,use_container_width=True,hide_index=True)
 st.info('fg_allocated_plan เป็นการจัดสรรในแผน ไม่ใช่การส่งจริง · MATERIAL_COVERED หมายถึงวัสดุตามแผนครอบคลุมเท่านั้น ไม่รับประกันกำลังเครื่องหรือวันส่ง')
with tabs[4]:
 st.subheader('เส้นทางการคำนวณวัสดุ')
 st.dataframe(requirements,hide_index=True,use_container_width=True)
 st.code('Good_new = max(0, Order_backlog - FG_allocated)\nStart = ceil(Good_new / (1 - ng_rate))\nGR = Start * BOM + Startup_loss\nProjected_stock(t) = Free_stock + Receipts_due(t) - Cumulative_GR(t)\nSS = z * sample_sd_daily * sqrt(lead_days)  # หรือค่า manual\nROP = mean_daily * lead_days + SS\nOrder_suggestion = max(0, SS - min_projected_stock)  # ปัด pack/MOQ\nLatest_order_at = first_below_SS_or_shortage - lead_days',language='text')
 st.caption('ng_rate เฉพาะ Scrap ระหว่าง Run ไม่รวม Startup ซ้ำ · ประวัติใช้จริงรายวันใช้คำนวณ SS/ROP ไม่ได้นำมาบวกเป็นความต้องการวัสดุซ้ำจาก Orders อีกครั้ง')
 st.warning('ขอบเขตรุ่น Excel นี้: ไม่มีระบบลง Ledger จริง/หลายคนพร้อมกัน ไม่มีการจัดคิวเครื่อง Changeover matrix หรือ New Model scenario อัตโนมัติ; ใช้ lead time สินค้าที่กรอกเป็นวันปฏิทินเพื่อกำหนดวันต้องใช้วัสดุ; Forecast หนึ่งเดือน ส่วน Orders อนาคตมีเท่าที่กรอก; หนึ่ง Order ที่ผลิตใหม่เป็นหนึ่งล็อต Startup')
with tabs[5]:
 st.subheader('ข้อมูลที่อ่านจาก Excel')
 sheet=st.selectbox('เลือกชีต',list(input_tabs));st.dataframe(input_tabs[sheet],hide_index=True,use_container_width=True)
 frames={'Forecast':forecasts,'Material_Plan':materials,'Order_Readiness':orders,'Requirements':requirements,'Timeline':timeline,'Delivery_KPI':pd.DataFrame([result['kpi']])}
 st.download_button('ดาวน์โหลดผลคำนวณเป็น Excel',safe_report(frames),file_name='GateCAT_forecast_and_material_plan.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
 st.caption('แอปอ่านและคำนวณจากไฟล์ที่อัปโหลด ไม่แก้ต้นฉบับ ไม่บันทึกลงเซิร์ฟเวอร์ถาวรในโค้ดนี้ และไม่ใช่ฐานข้อมูลหลายผู้ใช้ ข้อมูลจริงต้องได้รับอนุญาตและมีระบบสิทธิก่อนนำขึ้นเว็บสาธารณะ')
st.divider();st.caption('GateCAT · ต้นแบบอัปโหลด Excel และพยากรณ์อัตโนมัติ · ข้อมูลตัวอย่างสมมติ ไม่ใช่ผลดำเนินงานบริษัท')
