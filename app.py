import streamlit as st
import pandas as pd
import sqlite3, hashlib, io, re, os
from datetime import datetime

APP_DIR=os.path.dirname(os.path.abspath(__file__))
DB_PATH=os.path.join(APP_DIR,'tradeflow.db')
st.set_page_config(page_title='TradeFlow Analyzer', page_icon='📊', layout='wide', initial_sidebar_state='collapsed')
st.title('📊 Анализ')
st.caption('Excel/CSV импорт · Маҳаллий база · Қидирув · Фильтрлар · Ҳисобот · Экспорт')

ALIASES={
 'invoice_no':['фактура №','фактура номер','invoice number','invoice no','invoice','factura'],
 'invoice_date':['фактура санаси','фактура даты','invoice date','date'],
 'year':['йил','year'],
 'seller_id':['сотувчи стир','сотувчи стир/жшшири','сотувчи стир/жшшири','seller tin','seller id','seller stir','стир продавца'],
 'seller_name':['сотувчи номи','seller name','продавец'],
 'buyer_id':['сотиб олувчи стир','сотиб олувчи стир/жшшири','сотиб олувчи стир/жшшири','buyer tin','buyer id','buyer stir','покупатель стир'],
 'buyer_name':['сотиб олувчи номи','buyer name','покупатель'],
 'product':['маҳсулот ёки хизмат номи','товар номи','product name','товар','item'],
 'mxik':['товар хизматлар ягона электрон миллий каталоги бўйича идентификация коди','ягона электрон миллий каталоги бўйича идентификация коди','идентификация коди','мхик','mxik','product code'],
 'unit':['ўлчов бирлиги','unit','единица'], 'qty':['сони','quantity','qty','количество'],
 'price':['нархи','price','цена'], 'vat':['ққс (сумма)','ққс сумма','ққс','ккс','vat','ндс'],
 'total_vat':['етказиш баҳоси ққс билан','amount with vat','total','сумма с ндс'],
 'net':['ккc сиз киймати','ккс сиз киймати','net amount','amount without vat'],
 'direction':['изох','direction','тип','type']}

def norm(x):
 s=str(x).strip().lower().replace('\n',' ').replace('ё','е').replace('ў','у')
 s=s.replace('қ','к').replace('ғ','г')
 return re.sub(r'\s+',' ',s)
def detect_header(raw):
 best_i,best_score=0,-1
 for i in range(min(len(raw),20)):
  vals=[norm(v) for v in raw.iloc[i].tolist() if pd.notna(v)]
  score=sum(any(alias in v for alias_list in ALIASES.values() for alias in alias_list if len(alias)>4) for v in vals)
  score+=sum(any(k in v for k in ['фактура №','сотувчи номи','сотиб олувчи номи','маҳсулот ёки хизмат номи']) for v in vals)*3
  if score>best_score: best_i,best_score=i,score
 return best_i,best_score

def match_col(columns,key):
 aliases=[norm(a) for a in ALIASES[key]]
 for c in columns:
  n=norm(c)
  if n in aliases: return c
 for c in columns:
  n=norm(c)
  for a in aliases:
   if len(a)>=7 and (a in n or n in a): return c
 return None

def read_sheets(upload):
 upload.seek(0)
 if upload.name.lower().endswith('.csv'):
  raw=pd.read_csv(upload,header=None)
  h,score=detect_header(raw)
  raw.columns=raw.iloc[h]; df=raw.iloc[h+1:].copy(); return [('CSV',df,score)]
 xls=pd.ExcelFile(upload)
 result=[]
 for sheet in xls.sheet_names:
  # Summary sheet is not invoice-detail data; skip it explicitly.
  if norm(sheet) == norm('ТАХЛИЛ'):
   result.append((sheet,None,0)); continue
  raw=pd.read_excel(xls,sheet_name=sheet,header=None,nrows=30)
  if raw.empty: continue
  h,score=detect_header(raw)
  if score<3: result.append((sheet,None,score)); continue
  full=pd.read_excel(xls,sheet_name=sheet,header=h)
  full=full.dropna(how='all').copy()
  full.columns=[str(c).strip() for c in full.columns]
  result.append((sheet,full,score))
 return result

def clean(df,filename,sheet):
 cols={k:match_col(df.columns,k) for k in ALIASES}
 data=pd.DataFrame(index=df.index)
 for k,c in cols.items(): data[k]=df[c] if c is not None else pd.NA
 data['direction']=data['direction'].astype('string')
 if data.direction.isna().all():
  if 'кирим' in norm(sheet): data['direction']='incoming'
  elif 'чиким' in norm(sheet): data['direction']='outgoing'
  else: data['direction']='unspecified'
 for k in ['seller_id','buyer_id','mxik']:
  data[k]=data[k].astype('string').str.replace(r'\.0$','',regex=True).str.strip()
 for k in ['qty','price','vat','total_vat','net']:
  data[k]=pd.to_numeric(data[k].astype('string').str.replace(' ','').str.replace(',','.',regex=False),errors='coerce')
 if data.year.isna().any() and data.invoice_date.notna().any():
  parsed=pd.to_datetime(data.invoice_date,errors='coerce',dayfirst=True)
  data.year=data.year.astype('string').fillna(parsed.dt.year.astype('Int64').astype('string'))
 # Normalize Excel date objects before hashing or SQLite insertion.
 for c in ['invoice_date']:
  data[c]=data[c].map(lambda v: v.isoformat(sep=' ',timespec='seconds') if isinstance(v,(pd.Timestamp,datetime)) else v)
 data['source_file']=filename; data['source_sheet']=sheet; data['imported_at']=datetime.now().isoformat(timespec='seconds')
 keycols=['invoice_no','invoice_date','seller_id','buyer_id','product','mxik','qty','price','total_vat']
 data['fingerprint']=data[keycols].astype('string').fillna('').agg('|'.join,axis=1).map(lambda s:hashlib.sha256(s.encode()).hexdigest())
 # Drop empty rows that are not actual detail entries
 data=data[data[['invoice_no','seller_id','buyer_id','product']].notna().any(axis=1)]
 return data,cols

def initdb():
 with sqlite3.connect(DB_PATH) as con:
  con.execute('''CREATE TABLE IF NOT EXISTS lines (fingerprint TEXT UNIQUE, invoice_no TEXT, invoice_date TEXT, year TEXT, seller_id TEXT, seller_name TEXT, buyer_id TEXT, buyer_name TEXT, product TEXT, mxik TEXT, unit TEXT, qty REAL, price REAL, vat REAL, total_vat REAL, net REAL, direction TEXT, source_file TEXT, source_sheet TEXT, imported_at TEXT)''')
initdb()

with st.sidebar:
 st.header('Импорт ва маълумотлар')
 uploads=st.file_uploader('Excel ёки CSV файлни танланг',type=['xlsx','xls','csv'],accept_multiple_files=True)
 if uploads and st.button('Импортни бошлаш',type='primary',use_container_width=True):
  total_new=total_dup=0; logs=[]
  for file in uploads:
   try:
    sheets=read_sheets(file)
    for name,df,score in sheets:
     if df is None:
      logs.append(f'{file.name} / {name}: детал устунлари топилмади (умумий ҳисобот ёки мос келмаган варақ).'); continue
     cleaned,mapping=clean(df,file.name,name)
     essential=['invoice_no','seller_id','buyer_id','product']
     missing=[k for k in essential if mapping[k] is None]
     if len(missing)>=2 or cleaned.empty:
      logs.append(f'{file.name} / {name}: ўтказиб юборилди; мос устунлар етарли эмас.'); continue
     inserted=0; duplicates=0
     records=cleaned.to_dict('records')
     def sqlite_value(v):
      if pd.isna(v): return None
      if isinstance(v,(pd.Timestamp,datetime)): return v.isoformat(sep=' ',timespec='seconds')
      if hasattr(v,'item'):
       try: v=v.item()
       except Exception: pass
      if isinstance(v,(str,int,float,bool)) or v is None: return v
      return str(v)
     with sqlite3.connect(DB_PATH) as con:
      for r in records:
       fields=list(r.keys()); vals=[sqlite_value(r[k]) for k in fields]
       q=f"INSERT OR IGNORE INTO lines ({','.join(fields)}) VALUES ({','.join(['?']*len(fields))})"
       cur=con.execute(q,vals)
       inserted+=cur.rowcount==1; duplicates+=cur.rowcount==0
     total_new+=inserted; total_dup+=duplicates
     logs.append(f'{file.name} / {name}: {len(cleaned)} та детал сатри, янги {inserted}, такрорий {duplicates}.')
   except Exception as e: logs.append(f'{file.name}: хатолик — {type(e).__name__}: {e}')
  st.session_state['import_logs']=logs; st.session_state['import_summary']=(total_new,total_dup)
 if st.button('Базани тозалаш',help='Барча импорт қилинган ёзувларни ўчиради'):
  with sqlite3.connect(DB_PATH) as con: con.execute('DELETE FROM lines')
  st.session_state.pop('import_logs',None); st.rerun()

if 'import_logs' in st.session_state:
 new,dup=st.session_state.get('import_summary',(0,0))
 st.sidebar.success(f'Импорт натижаси: янги {new:,}; такрорий {dup:,}')
 with st.sidebar.expander('Импорт тафсилотлари',expanded=True):
  for line in st.session_state.import_logs: st.write(line)

with sqlite3.connect(DB_PATH) as con: data=pd.read_sql_query('SELECT rowid AS id,* FROM lines',con)
if data.empty:
 st.info('Excel/CSV файлни юкланг ва чап панелдаги **Импортни бошлаш** тугмасини босинг.')
 st.markdown('### Нима учун жадвал ҳозирча бўш?\nФайлни танлашнинг ўзи импорт қилмайди. Импорт тугмасини босгандан кейин варақлар ўқилади. Импорт тафсилотларида қайси варақ ўқилгани ёки ўтказиб юборилганини кўрасиз.')
 st.stop()

st.subheader('Фильтрлар')
a,b,c,d=st.columns(4)
years=sorted(data['year'].dropna().astype(str).unique().tolist()); year_filter=a.multiselect('Йил',years,default=years)
# Build one searchable company selector whose option label contains both STIR and name.
company_names={}
for _, row in data.iterrows():
 for id_col,name_col in [('seller_id','seller_name'),('buyer_id','buyer_name')]:
  cid=row[id_col]
  if pd.notna(cid) and str(cid).strip():
   cid=str(cid).strip()
   cname=row[name_col]
   if pd.notna(cname) and str(cname).strip():
    company_names.setdefault(cid,set()).add(str(cname).strip())
company_ids=sorted(company_names)
def company_label(cid):
 names=' / '.join(sorted(company_names.get(str(cid),set())))
 return f'{cid} — {names}' if names else str(cid)
company_filter=b.multiselect('СТИР / корхона номи',company_ids,format_func=company_label,placeholder='СТИР ёки корхона номини қидиринг')
products=sorted(data['product'].dropna().astype(str).unique().tolist()); product_filter=c.multiselect('Маҳсулот',products,placeholder='Маҳсулотни қидиринг')
dirs=sorted(data['direction'].dropna().astype(str).unique().tolist())
dir_labels={'incoming':'Кирим','outgoing':'Чиқим','unspecified':'Йўналиш аниқланмаган'}
direction_filter=d.multiselect('Ҳаракат йўналиши',dirs,format_func=lambda x:dir_labels.get(x,str(x)))
filtered=data.copy()
if year_filter: filtered=filtered[filtered['year'].astype(str).isin(year_filter)]
if company_filter: filtered=filtered[filtered['seller_id'].astype(str).isin(company_filter)|filtered['buyer_id'].astype(str).isin(company_filter)]
if product_filter: filtered=filtered[filtered['product'].astype(str).isin(product_filter)]
if direction_filter: filtered=filtered[filtered['direction'].astype(str).isin(direction_filter)]

st.subheader('Умумий кўрсаткичлар')
col1,col2,col3,col4=st.columns(4)
col1.metric('Сатрлар сони',f'{len(filtered):,}'); col2.metric('Инвойслар сони',f'{filtered["invoice_no"].nunique():,}'); col3.metric('Жами миқдор',f'{filtered["qty"].sum():,.2f}'); col4.metric('Жами сумма',f'{filtered["total_vat"].sum():,.2f}')
tabs=st.tabs(['Ёзувлар','Корхоналар','Маҳсулотлар','Импорт ва экспорт'])
with tabs[0]:
 # Requested columns appear first; remaining useful fields follow.
 display_order=['invoice_date','product','seller_name','seller_id','buyer_id','buyer_name','invoice_no','year','mxik','unit','qty','price','vat','total_vat','net','direction','source_file','source_sheet']
 display_labels={'invoice_date':'Инвойс санаси','product':'Маҳсулот номи','seller_name':'Сотувчи номи','seller_id':'Сотувчи СТИР','buyer_id':'Харидор СТИР','buyer_name':'Харидор номи','invoice_no':'Инвойс рақами','year':'Йил','mxik':'МХИК коди','unit':'Ўлчов бирлиги','qty':'Миқдор','price':'Нарх','vat':'ҚҚС','total_vat':'ҚҚС билан сумма','net':'ҚҚСсиз сумма','direction':'Йўналиш','source_file':'Манба файл','source_sheet':'Манба варақ'}
 visible_cols=[x for x in display_order if x in filtered.columns]
 table=filtered[visible_cols].rename(columns=display_labels)
 if 'Йўналиш' in table.columns: table['Йўналиш']=table['Йўналиш'].map(dir_labels).fillna(table['Йўналиш'])
 st.dataframe(table,hide_index=True,use_container_width=True)
with tabs[1]:
 s=filtered.groupby(['seller_id','seller_name'],dropna=False).agg(sotuvchi_satr=('id','count'),sotuvchi_summa=('total_vat','sum')).reset_index().rename(columns={'seller_id':'СТИР','seller_name':'Сотувчи номи','sotuvchi_satr':'Сотувчи ёзувлари','sotuvchi_summa':'Сотувчи суммаси'})
 b=filtered.groupby(['buyer_id','buyer_name'],dropna=False).agg(xaridor_satr=('id','count'),xaridor_summa=('total_vat','sum')).reset_index().rename(columns={'buyer_id':'СТИР','buyer_name':'Харидор номи','xaridor_satr':'Харидор ёзувлари','xaridor_summa':'Харидор суммаси'})
 st.dataframe(pd.concat([s,b],ignore_index=True),use_container_width=True,hide_index=True)
with tabs[2]:
 p=filtered.groupby(['mxik','product','unit'],dropna=False).agg(satr=('id','count'),miqdor=('qty','sum'),summa=('total_vat','sum')).reset_index().sort_values('summa',ascending=False).rename(columns={'mxik':'МХИК коди','product':'Маҳсулот номи','unit':'Ўлчов бирлиги','satr':'Сатрлар сони','miqdor':'Жами миқдор','summa':'Жами сумма'})
 st.dataframe(p,use_container_width=True,hide_index=True)
with tabs[3]:
 st.write('Базага импорт қилинган манба файллар:')
 sources=data.groupby(['source_file','source_sheet'],dropna=False).agg(satr=('id','count'),invoice=('invoice_no','nunique')).reset_index().rename(columns={'source_file':'Манба файл','source_sheet':'Манба варақ','satr':'Сатрлар сони','invoice':'Инвойслар сони'})
 st.dataframe(sources,use_container_width=True,hide_index=True)
 csv=filtered.to_csv(index=False).encode('utf-8-sig'); st.download_button('CSV файлни экспорт қилиш',csv,'tradeflow_export.csv','text/csv')
 out=io.BytesIO()
 with pd.ExcelWriter(out,engine='openpyxl') as writer:
  filtered.to_excel(writer,index=False,sheet_name='Маълумотлар')
  filtered.groupby(['mxik','product','unit'],dropna=False).agg(qty=('qty','sum'),amount=('total_vat','sum')).reset_index().to_excel(writer,index=False,sheet_name='Маҳсулотлар')
 st.download_button('Excel файлни экспорт қилиш',out.getvalue(),'tradeflow_export.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
st.caption('Маҳаллий таҳлил прототипи. Натижалар фақат импорт қилинган маълумотларни умумлаштиради ва ўз-ўзидан ҳуқуқбузарлик ҳақида хулоса бермайди.')
