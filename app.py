import os
import io
import pandas as pd
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pymongo import MongoClient
from datetime import datetime
import re

app = FastAPI()

# 允许跨域
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ---------- 连接MongoDB ----------
MONGO_URI = os.getenv("MONGO_URI")
client = MongoClient(MONGO_URI)
db = client["sales_dashboard"]
collection = db["daily_sales"]

# ---------- 智能解析Excel ----------
def parse_excel(df):
    """自动识别列名，提取日期、销量、销售额"""
    # 查找日期列（包含'日期'、'时间'、'date'等字样）
    date_col = None
    for col in df.columns:
        if any(keyword in str(col).lower() for keyword in ['日期', '时间', 'date', '业务日期', '单据日期']):
            date_col = col
            break
    if date_col is None:
        # 如果找不到，尝试第二列（很多金蝶导出日期在第二列）
        if len(df.columns) >= 2:
            date_col = df.columns[1]
        else:
            raise ValueError("无法找到日期列，请确保Excel包含日期信息")
    
    # 查找数量列（包含'数量'、'销量'、'出库数量'等）
    qty_col = None
    for col in df.columns:
        if any(keyword in str(col).lower() for keyword in ['数量', '销量', '出库数量', 'qty']):
            qty_col = col
            break
    if qty_col is None:
        # 尝试第5列（常见位置）
        if len(df.columns) >= 5:
            qty_col = df.columns[4]
        else:
            raise ValueError("无法找到数量列")
    
    # 查找金额列（包含'金额'、'销售额'、'价税合计'等）
    amt_col = None
    for col in df.columns:
        if any(keyword in str(col).lower() for keyword in ['金额', '销售额', '价税合计', 'amount']):
            amt_col = col
            break
    if amt_col is None:
        # 尝试第6列
        if len(df.columns) >= 6:
            amt_col = df.columns[5]
        else:
            amt_col = qty_col  # 如果没有金额列，就用数量代替（避免报错）
    
    # 提取数据
    df_clean = df[[date_col, qty_col, amt_col]].copy()
    df_clean.columns = ["日期", "销量", "销售额"]
    
    # 智能解析日期（尝试多种格式）
    def parse_date(val):
        if pd.isna(val):
            return pd.NaT
        if isinstance(val, (pd.Timestamp, datetime)):
            return val
        # 如果是类似 "20260906" 的纯数字
        if isinstance(val, (int, float)):
            val = str(int(val))
        # 尝试常见格式
        for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%Y%m%d", "%y-%m-%d", "%d-%m-%Y", "%m-%d-%Y"]:
            try:
                return pd.to_datetime(val, format=fmt)
            except:
                continue
        try:
            return pd.to_datetime(val, errors='coerce')
        except:
            return pd.NaT
    
    df_clean["日期"] = df_clean["日期"].apply(parse_date)
    df_clean = df_clean.dropna(subset=["日期"])
    
    # 确保数值类型
    df_clean["销量"] = pd.to_numeric(df_clean["销量"], errors='coerce').fillna(0)
    df_clean["销售额"] = pd.to_numeric(df_clean["销售额"], errors='coerce').fillna(0)
    
    # 按日汇总
    daily = df_clean.groupby(df_clean["日期"].dt.date).agg({
        "销量": "sum",
        "销售额": "sum"
    }).reset_index()
    daily.columns = ["日期", "总销量", "总销售额"]
    return daily

# ---------- 上传Excel ----------
@app.post("/upload")
async def upload_excel(file: UploadFile = File(...)):
    try:
        content = await file.read()
        df = pd.read_excel(io.BytesIO(content))
        daily = parse_excel(df)
        
        # 存入MongoDB
        collection.delete_many({})
        records = daily.to_dict("records")
        for r in records:
            r["日期"] = r["日期"].strftime("%Y-%m-%d")
        collection.insert_many(records)
        
        return JSONResponse({
            "success": True,
            "message": f"成功上传 {len(records)} 天数据",
            "count": len(records)
        })
    except Exception as e:
        return JSONResponse({
            "success": False,
            "message": str(e)
        }, status_code=400)

# ---------- API ----------
@app.get("/api/data")
def get_data():
    data = list(collection.find({}, {"_id": 0}))
    return JSONResponse(data)

@app.get("/api/summary")
def get_summary():
    data = list(collection.find({}, {"_id": 0}))
    if not data:
        return JSONResponse({"今日销量": 0, "本月销量": 0, "本年销量": 0, "今日销售额": 0, "本月销售额": 0, "本年销售额": 0})
    df = pd.DataFrame(data)
    df["日期"] = pd.to_datetime(df["日期"])
    today = pd.Timestamp.now().date()
    return JSONResponse({
        "今日销量": int(df[df["日期"].dt.date == today]["总销量"].sum()),
        "本月销量": int(df[df["日期"].dt.month == today.month]["总销量"].sum()),
        "本年销量": int(df[df["日期"].dt.year == today.year]["总销量"].sum()),
        "今日销售额": int(df[df["日期"].dt.date == today]["总销售额"].sum()),
        "本月销售额": int(df[df["日期"].dt.month == today.month]["总销售额"].sum()),
        "本年销售额": int(df[df["日期"].dt.year == today.year]["总销售额"].sum()),
    })

# ---------- 大屏HTML ----------
HTML_PAGE = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=1920, initial-scale=1.0">
    <title>产销量数据大屏</title>
    <script src="https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js"></script>
    <style>
        *{margin:0;padding:0;box-sizing:border-box}
        body{background:radial-gradient(ellipse at center,#0a0e27,#030514);color:#e0e0e0;font-family:'Microsoft YaHei',Arial,sans-serif;padding:30px;min-height:100vh}
        .dashboard{max-width:1920px;margin:0 auto}
        .header{text-align:center;padding:20px 0 30px;border-bottom:2px solid rgba(74,95,193,0.3);margin-bottom:30px}
        .header h1{font-size:48px;font-weight:700;background:linear-gradient(to right,#4a5fc1,#7986cb,#4a5fc1);-webkit-background-clip:text;-webkit-text-fill-color:transparent;letter-spacing:8px}
        .header .sub{color:#7986cb;font-size:18px;letter-spacing:4px;margin-top:8px;opacity:0.7}
        .upload-area{background:rgba(20,22,40,0.8);border:2px dashed rgba(121,134,203,0.3);border-radius:16px;padding:30px 40px;margin-bottom:30px;display:flex;align-items:center;justify-content:space-between;backdrop-filter:blur(10px);transition:border-color 0.3s}
        .upload-area:hover{border-color:rgba(121,134,203,0.6)}
        .upload-area .left{display:flex;align-items:center;gap:20px}
        .upload-area .left .icon{font-size:40px;color:#7986cb}
        .upload-area .left .info .title{color:#e0e0e0;font-size:18px;font-weight:600}
        .upload-area .left .info .desc{color:#7986cb;font-size:14px;margin-top:4px}
        .upload-area .right{display:flex;align-items:center;gap:15px}
        .file-btn{background:linear-gradient(135deg,#4a5fc1,#7986cb);color:#fff;border:none;padding:12px 32px;border-radius:10px;font-size:16px;cursor:pointer;transition:all 0.3s;font-weight:600}
        .file-btn:hover{transform:translateY(-2px);box-shadow:0 8px 30px rgba(74,95,193,0.4)}
        .file-input{display:none}
        .upload-btn{background:linear-gradient(135deg,#ffd54f,#ffb300);color:#1a1a2e;border:none;padding:12px 40px;border-radius:10px;font-size:16px;cursor:pointer;transition:all 0.3s;font-weight:700}
        .upload-btn:hover{transform:translateY(-2px);box-shadow:0 8px 30px rgba(255,213,79,0.4)}
        .upload-btn:disabled{opacity:0.4;cursor:not-allowed;transform:none}
        .file-name{color:#4fc3f7;font-size:14px;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
        .upload-status{margin-top:12px;text-align:center;font-size:15px}
        .upload-status.success{color:#81c784}
        .upload-status.error{color:#ef5350}
        .upload-status.loading{color:#ffd54f}
        .kpi-row{display:grid;grid-template-columns:repeat(6,1fr);gap:20px;margin-bottom:30px}
        .kpi-card{background:rgba(20,22,40,0.85);border-radius:16px;padding:24px 20px;text-align:center;border:1px solid rgba(121,134,203,0.25);box-shadow:0 8px 32px rgba(0,0,0,0.5);backdrop-filter:blur(10px);transition:transform 0.3s}
        .kpi-card:hover{transform:translateY(-4px)}
        .kpi-card .label{font-size:16px;color:#7986cb;letter-spacing:2px}
        .kpi-card .value{font-size:38px;font-weight:700;color:#fff;margin-top:10px}
        .kpi-card .value.gold{color:#ffd54f}
        .kpi-card .value.blue{color:#4fc3f7}
        .kpi-card .value.green{color:#81c784}
        .kpi-card .value.pink{color:#f48fb1}
        .kpi-card .value.purple{color:#ce93d8}
        .kpi-card .value.lightblue{color:#90caf9}
        .chart-grid{display:grid;grid-template-columns:1fr 1fr;gap:25px;margin-bottom:25px}
        .chart-box{background:rgba(20,22,40,0.8);border-radius:16px;padding:20px;border:1px solid rgba(121,134,203,0.15);box-shadow:0 8px 32px rgba(0,0,0,0.4)}
        .chart-box.full{grid-column:1/-1}
        .chart-box .title{font-size:16px;color:#7986cb;padding-bottom:12px;border-bottom:1px solid rgba(121,134,203,0.1);margin-bottom:12px;letter-spacing:2px}
        .chart-container{width:100%;height:320px}
        .footer{text-align:center;padding:20px;color:#3a4a7a;font-size:14px;border-top:1px solid rgba(121,134,203,0.1);margin-top:20px}
    </style>
</head>
<body>
<div class="dashboard">
    <div class="header"><h1>📊 企业产销量数据大屏</h1><div class="sub">金蝶EAS数据驱动 · 上传即更新</div></div>
    <div class="upload-area">
        <div class="left">
            <div class="icon">📁</div>
            <div class="info">
                <div class="title">上传销售数据</div>
                <div class="desc">支持 .xlsx 格式 · 自动识别日期、销量、金额</div>
            </div>
        </div>
        <div class="right">
            <span class="file-name" id="fileName">未选择文件</span>
            <button class="file-btn" onclick="document.getElementById('fileInput').click()">选择文件</button>
            <input type="file" id="fileInput" class="file-input" accept=".xlsx">
            <button class="upload-btn" id="uploadBtn" onclick="uploadFile()">上传到系统</button>
        </div>
    </div>
    <div id="uploadStatus" class="upload-status"></div>
    <div class="kpi-row">
        <div class="kpi-card"><div class="label">📅 今日销量</div><div class="value gold" id="k1">--</div></div>
        <div class="kpi-card"><div class="label">📆 本月销量</div><div class="value blue" id="k2">--</div></div>
        <div class="kpi-card"><div class="label">📈 本年销量</div><div class="value green" id="k3">--</div></div>
        <div class="kpi-card"><div class="label">💰 今日销售额</div><div class="value pink" id="k4">--</div></div>
        <div class="kpi-card"><div class="label">💰 本月销售额</div><div class="value purple" id="k5">--</div></div>
        <div class="kpi-card"><div class="label">💰 本年销售额</div><div class="value lightblue" id="k6">--</div></div>
    </div>
    <div class="chart-grid">
        <div class="chart-box full"><div class="title">📈 日销量趋势</div><div class="chart-container" id="trendChart"></div></div>
    </div>
    <div class="chart-grid">
        <div class="chart-box"><div class="title">📊 月度销量</div><div class="chart-container" id="monthChart"></div></div>
        <div class="chart-box"><div class="title">📊 年度销量</div><div class="chart-container" id="yearChart"></div></div>
    </div>
    <div class="footer">数据来源：金蝶EAS · 上传后自动更新</div>
</div>
<script>
document.getElementById('fileInput').addEventListener('change', function(e){
    const name = e.target.files[0] ? e.target.files[0].name : '未选择文件';
    document.getElementById('fileName').textContent = name;
    document.getElementById('uploadStatus').textContent = '';
});
async function uploadFile(){
    const input = document.getElementById('fileInput');
    const btn = document.getElementById('uploadBtn');
    const status = document.getElementById('uploadStatus');
    if(!input.files || input.files.length === 0){
        status.className = 'upload-status error';
        status.textContent = '⚠️ 请先选择一个Excel文件';
        return;
    }
    const file = input.files[0];
    const formData = new FormData();
    formData.append('file', file);
    btn.disabled = true;
    btn.textContent = '上传中...';
    status.className = 'upload-status loading';
    status.textContent = '⏳ 正在上传并解析数据...';
    try{
        const resp = await fetch('/upload', {
            method: 'POST',
            body: formData
        });
        const result = await resp.json();
        if(result.success){
            status.className = 'upload-status success';
            status.textContent = '✅ ' + result.message + '，大屏已自动更新！';
            fetchData();
        }else{
            status.className = 'upload-status error';
            status.textContent = '❌ 上传失败：' + result.message;
        }
    }catch(e){
        status.className = 'upload-status error';
        status.textContent = '❌ 网络错误：' + e.message;
    }finally{
        btn.disabled = false;
        btn.textContent = '上传到系统';
    }
}
function fetchData(){
    fetch('/api/data').then(r=>r.json()).then(data=>{
        if(!data || data.length === 0){
            const t = echarts.init(document.getElementById('trendChart'),'dark');
            t.setOption({title:{text:'暂无数据，请上传Excel',textStyle:{color:'#666',fontSize:20},left:'center',top:'center'}});
            return;
        }
        renderCharts(data);
    });
    fetch('/api/summary').then(r=>r.json()).then(d=>{
        document.getElementById('k1').textContent = d.今日销量.toLocaleString();
        document.getElementById('k2').textContent = d.本月销量.toLocaleString();
        document.getElementById('k3').textContent = d.本年销量.toLocaleString();
        document.getElementById('k4').textContent = d.今日销售额.toLocaleString();
        document.getElementById('k5').textContent = d.本月销售额.toLocaleString();
        document.getElementById('k6').textContent = d.本年销售额.toLocaleString();
    });
}
function renderCharts(data){
    const dates = data.map(d=>d.日期);
    const sales = data.map(d=>d.总销量);
    const amts = data.map(d=>d.总销售额);
    const t = echarts.init(document.getElementById('trendChart'),'dark');
    t.setOption({
        tooltip:{trigger:'axis'},
        legend:{data:['销量','销售额'],textStyle:{color:'#aaa'}},
        xAxis:{type:'category',data:dates,axisLabel:{color:'#888',fontSize:12}},
        yAxis:[
            {type:'value',name:'销量',nameTextStyle:{color:'#888'},axisLabel:{color:'#888'}},
            {type:'value',name:'销售额',nameTextStyle:{color:'#888'},axisLabel:{color:'#888'}}
        ],
        series:[
            {name:'销量',type:'line',data:sales,smooth:true,lineStyle:{color:'#ffd54f',width:3},areaStyle:{color:'rgba(255,213,79,0.15)'},symbol:'circle',symbolSize:6},
            {name:'销售额',type:'line',data:amts,smooth:true,yAxisIndex:1,lineStyle:{color:'#4fc3f7',width:3},areaStyle:{color:'rgba(79,195,247,0.15)'},symbol:'circle',symbolSize:6}
        ],
        grid:{top:30,bottom:30,left:60,right:60}
    });
    window.addEventListener('resize', ()=>t.resize());
    const last6 = data.slice(-6);
    const m = echarts.init(document.getElementById('monthChart'),'dark');
    m.setOption({
        tooltip:{trigger:'axis'},
        xAxis:{type:'category',data:last6.map(d=>d.日期.slice(0,7)),axisLabel:{color:'#888'}},
        yAxis:{type:'value',axisLabel:{color:'#888'}},
        series:[{type:'bar',data:last6.map(d=>d.总销量),itemStyle:{color:'#4a5fc1',borderRadius:[4,4,0,0]}}],
        grid:{top:20,bottom:30,left:50,right:20}
    });
    window.addEventListener('resize', ()=>m.resize());
    const y = {};
    data.forEach(d=>{const yy=d.日期.slice(0,4); y[yy]=(y[yy]||0)+d.总销量});
    const yc = echarts.init(document.getElementById('yearChart'),'dark');
    yc.setOption({
        tooltip:{trigger:'axis'},
        xAxis:{type:'category',data:Object.keys(y),axisLabel:{color:'#888'}},
        yAxis:{type:'value',axisLabel:{color:'#888'}},
        series:[{type:'bar',data:Object.values(y),itemStyle:{color:'#7986cb',borderRadius:[4,4,0,0]}}],
        grid:{top:20,bottom:30,left:50,right:20}
    });
    window.addEventListener('resize', ()=>yc.resize());
}
fetchData();
</script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
def dashboard():
    return HTMLResponse(HTML_PAGE)