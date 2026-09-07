import os
import io
import re
import pandas as pd

from datetime import datetime

from fastapi import FastAPI, UploadFile, File
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

from pymongo import MongoClient


# =====================================================
# FastAPI
# =====================================================

app = FastAPI(
    title="西瑞集团经营数据驾驶舱",
    version="3.0 (优化版)"
)


app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)



# =====================================================
# MongoDB
# =====================================================

MONGO_URI = os.getenv(
    "MONGO_URI",
    "mongodb://localhost:27017"
)


client = MongoClient(MONGO_URI)

db = client["sales_dashboard"]

collection = db["daily_sales"]

raw_collection = db["raw_sales"]



# =====================================================
# 智能列识别（升级版）
# =====================================================

def smart_find_column(columns, keywords, fallback_index=None):
    """
    智能查找列：按关键词（精确/正则）匹配，最后按位置兜底
    """
    # 第一轮：精确包含关键词
    for col in columns:
        col_lower = str(col).lower().strip()
        for kw in keywords:
            if kw.lower() in col_lower:
                return col

    # 第二轮：正则匹配（如“出库数量”匹配“数量”）
    # 这里内置一些常见的模式映射
    pattern_map = {
        '日期': r'(日期|时间|date|业务日期|单据日期|下单日期)',
        '数量': r'(数量|销量|出库数量|销售数量|qty|吨)',
        '金额': r'(金额|销售额|价税合计|含税金额|amount|总计)',
        '产品': r'(产品|物料|商品|品种|名称)',
        '客户': r'(客户|购货单位|单位|往来单位|客户名称)'
    }
    # 如果 keywords 包含以上关键词之一，使用对应的正则
    for kw in keywords:
        for key, pattern in pattern_map.items():
            if kw.lower() in key or key in kw.lower():
                for col in columns:
                    if re.search(pattern, str(col), re.IGNORECASE):
                        return col
                break

    # 第三轮：按位置兜底
    if fallback_index is not None and fallback_index < len(columns):
        return columns[fallback_index]

    return None


# =====================================================
# 数据质量检查
# =====================================================

def validate_data(df):
    """
    检查数据质量，返回问题列表
    """
    issues = []

    # 1. 空值检查
    null_counts = df.isnull().sum()
    for col, count in null_counts.items():
        if count > 0:
            issues.append(f"⚠️ {col} 列有 {count} 行空值")

    # 2. 负值检查
    if '销量' in df.columns:
        neg = (df['销量'] < 0).sum()
        if neg > 0:
            issues.append(f"⚠️ 有 {neg} 行销量为负数")

    if '销售额' in df.columns:
        neg = (df['销售额'] < 0).sum()
        if neg > 0:
            issues.append(f"⚠️ 有 {neg} 行销售额为负数")

    # 3. 异常值（3倍标准差）
    if '销量' in df.columns and len(df) > 1:
        mean = df['销量'].mean()
        std = df['销量'].std()
        if std > 0:
            outliers = ((df['销量'] - mean).abs() > 3 * std).sum()
            if outliers > 0:
                issues.append(f"⚠️ 有 {outliers} 行销量异常（超出3倍标准差）")

    # 4. 日期范围
    if '日期' in df.columns:
        min_date = df['日期'].min()
        max_date = df['日期'].max()
        issues.append(f"📅 数据日期范围：{min_date.date()} 至 {max_date.date()}")

    return issues


# =====================================================
# Excel解析核心（流式分块读取 + 智能识别）
# =====================================================

def parse_excel_streaming(file_content):
    """
    分块流式读取 Excel，逐块清洗，最后合并
    """
    all_chunks = []

    # 先读取第一块（或前几行）用于识别列位置
    first_chunk = pd.read_excel(
        io.BytesIO(file_content),
        nrows=5,
        header=0
    )
    columns = first_chunk.columns

    # 智能识别列
    date_col = smart_find_column(columns, ["日期", "时间", "date", "业务日期"], 0)
    qty_col = smart_find_column(columns, ["数量", "销量", "出库数量", "qty"], 4)
    amt_col = smart_find_column(columns, ["金额", "销售额", "价税合计", "amount"], 5)
    product_col = smart_find_column(columns, ["产品", "物料", "商品"], None)
    customer_col = smart_find_column(columns, ["客户", "购货单位", "单位"], None)

    # 确定要读取的列（按位置）
    # 由于分块读取时只能按索引或列名，我们使用列名列表
    cols_to_read = [date_col, qty_col, amt_col]
    if product_col:
        cols_to_read.append(product_col)
    if customer_col:
        cols_to_read.append(customer_col)

    # 分块读取
    for chunk in pd.read_excel(
        io.BytesIO(file_content),
        sheet_name=0,
        usecols=cols_to_read,
        chunksize=10000,
        dtype_backend='pyarrow'  # 加快速度，减少内存
    ):
        # 重命名列
        rename_map = {
            date_col: "日期",
            qty_col: "销量",
            amt_col: "销售额"
        }
        if product_col:
            rename_map[product_col] = "产品"
        if customer_col:
            rename_map[customer_col] = "客户"
        chunk = chunk.rename(columns=rename_map)

        # 转换数据类型
        chunk["日期"] = pd.to_datetime(chunk["日期"], errors='coerce')
        chunk["销量"] = pd.to_numeric(chunk["销量"], errors='coerce').fillna(0)
        chunk["销售额"] = pd.to_numeric(chunk["销售额"], errors='coerce').fillna(0)

        # 若没有产品/客户列，填充默认值
        if product_col is None:
            chunk["产品"] = "未知产品"
        if customer_col is None:
            chunk["客户"] = "未知客户"

        # 删除日期无效的行
        chunk = chunk.dropna(subset=["日期"])

        all_chunks.append(chunk)

    if not all_chunks:
        raise ValueError("未读取到有效数据")

    # 合并所有块
    data = pd.concat(all_chunks, ignore_index=True)
    return data


# =====================================================
# 上传Excel接口（优化版）
# =====================================================

@app.post("/upload")
async def upload_excel(
    file: UploadFile = File(...)
):
    try:
        content = await file.read()

        # 流式解析
        data = parse_excel_streaming(content)

        # 数据质量检查
        issues = validate_data(data)

        # 保存原始数据（明细）
        raw_collection.delete_many({})
        raw_collection.insert_many(
            data.to_dict("records")
        )

        # 按日汇总
        daily = (
            data
            .groupby(data["日期"].dt.date)
            .agg({
                "销量": "sum",
                "销售额": "sum"
            })
            .reset_index()
        )
        daily.columns = ["日期", "总销量", "总销售额"]

        collection.delete_many({})

        records = daily.to_dict("records")
        for r in records:
            r["日期"] = str(r["日期"])

        collection.insert_many(records)

        return {
            "success": True,
            "message": f"成功导入 {len(records)} 天数据",
            "total_rows": len(data),
            "issues": issues
        }

    except Exception as e:
        return JSONResponse(
            {
                "success": False,
                "message": str(e)
            },
            status_code=400
        )


# =====================================================
# 基础数据接口
# =====================================================

@app.get("/api/data")
def get_data():
    data = list(
        collection.find(
            {},
            {"_id": 0}
        )
    )
    return data


# =====================================================
# KPI接口
# =====================================================

@app.get("/api/summary")
def summary():
    data = list(
        collection.find(
            {},
            {"_id": 0}
        )
    )
    if not data:
        return {}

    df = pd.DataFrame(data)
    df["日期"] = pd.to_datetime(df["日期"])
    now = pd.Timestamp.now()

    return {
        "今日销量": int(df[df["日期"].dt.date == now.date()]["总销量"].sum()),
        "本月销量": int(df[df["日期"].dt.month == now.month]["总销量"].sum()),
        "本年销量": int(df[df["日期"].dt.year == now.year]["总销量"].sum()),
        "今日销售额": float(df[df["日期"].dt.date == now.date()]["总销售额"].sum()),
        "本月销售额": float(df[df["日期"].dt.month == now.month]["总销售额"].sum()),
        "本年销售额": float(df[df["日期"].dt.year == now.year]["总销售额"].sum()),
    }


# =====================================================
# 产品排行
# =====================================================

@app.get("/api/product_rank")
def product_rank():
    data = list(
        raw_collection.find(
            {},
            {"_id": 0}
        )
    )
    if not data:
        return []
    df = pd.DataFrame(data)
    result = (
        df.groupby("产品")["销量"]
        .sum()
        .sort_values(ascending=False)
        .head(10)
    )
    return [
        {"name": k, "value": float(v)}
        for k, v in result.items()
    ]


# =====================================================
# 客户排行
# =====================================================

@app.get("/api/customer_rank")
def customer_rank():
    data = list(
        raw_collection.find(
            {},
            {"_id": 0}
        )
    )
    if not data:
        return []
    df = pd.DataFrame(data)
    result = (
        df.groupby("客户")["销售额"]
        .sum()
        .sort_values(ascending=False)
        .head(10)
    )
    return [
        {"name": k, "value": float(v)}
        for k, v in result.items()
    ]


# =====================================================
# 自动化分析报告 API
# =====================================================

@app.get("/api/report")
def generate_report():
    data = list(collection.find({}, {"_id": 0}))
    if not data:
        return {"error": "暂无数据"}

    df = pd.DataFrame(data)
    df["日期"] = pd.to_datetime(df["日期"])
    df_sorted = df.sort_values("日期")

    # 基础统计
    report = {
        "总销量": float(df["总销量"].sum()),
        "总销售额": float(df["总销售额"].sum()),
        "日均销量": float(df["总销量"].mean()),
        "数据天数": len(df),
        "最高单日销量": float(df["总销量"].max()),
        "最高单日销量日期": df[df["总销量"] == df["总销量"].max()]["日期"].iloc[0].strftime("%Y-%m-%d"),
        "最低单日销量": float(df["总销量"].min()),
        "最低单日销量日期": df[df["总销量"] == df["总销量"].min()]["日期"].iloc[0].strftime("%Y-%m-%d"),
    }

    # 月度汇总
    monthly = df.groupby(df["日期"].dt.to_period("M"))["总销量"].sum()
    report["月度销量"] = {str(k): float(v) for k, v in monthly.items()}

    # 环比增长率
    if len(monthly) >= 2:
        last = monthly.iloc[-1]
        prev = monthly.iloc[-2]
        report["环比增长率"] = f"{((last - prev) / prev * 100):.1f}%"
        report["本月销量"] = float(last)
        report["上月销量"] = float(prev)

    return report


# =====================================================
# 企业驾驶舱HTML（保持不变，使用修复过的版本）
# =====================================================

HTML_PAGE = r"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width,initial-scale=1.0">
    <title>西瑞集团经营数据驾驶舱</title>
    <script src="https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js"></script>
    <style>
        *{margin:0;padding:0;box-sizing:border-box}
        body{background:linear-gradient(135deg,#071426,#0d2340);font-family:"Microsoft YaHei",Arial,sans-serif;color:white;min-height:100vh}
        .dashboard{width:95%;max-width:1920px;margin:auto;padding:25px}
        .header{height:100px;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid rgba(255,255,255,.15)}
        .title{font-size:38px;font-weight:bold;letter-spacing:6px;color:#f3c65a}
        .subtitle{font-size:16px;color:#9fb5d1;margin-top:10px}
        .time{font-size:18px;color:#9fd5ff}
        .upload-panel{margin-top:25px;background:rgba(255,255,255,.06);border-radius:15px;padding:20px;display:flex;justify-content:space-between;align-items:center}
        .upload-btn{background:#d8a43a;border:none;padding:12px 35px;border-radius:8px;font-size:16px;cursor:pointer;font-weight:bold}
        .upload-btn:hover{opacity:.85}
        .file-name{margin-right:20px;color:#9fd5ff}
        .kpi-grid{margin-top:25px;display:grid;grid-template-columns:repeat(6,1fr);gap:20px}
        .kpi-card{background:linear-gradient(145deg,rgba(255,255,255,.10),rgba(255,255,255,.03));border-radius:16px;padding:25px;text-align:center;border:1px solid rgba(255,255,255,.12)}
        .kpi-label{font-size:16px;color:#9fb5d1}
        .kpi-value{margin-top:15px;font-size:38px;font-weight:bold;color:#ffffff}
        .unit{font-size:15px;color:#d8a43a}
        .chart-grid{margin-top:25px;display:grid;grid-template-columns:2fr 1fr;gap:25px}
        .chart-box{background:rgba(255,255,255,.06);border-radius:15px;padding:20px;border:1px solid rgba(255,255,255,.12)}
        .chart-title{font-size:18px;color:#f3c65a;margin-bottom:15px}
        .chart{height:380px;width:100%}
        .small-chart{height:300px}
        .bottom-grid{margin-top:25px;display:grid;grid-template-columns:1fr 1fr;gap:25px}
        .rank-box{background:rgba(255,255,255,.06);border-radius:15px;padding:20px}
        .rank-item{display:flex;justify-content:space-between;padding:12px 0;border-bottom:1px solid rgba(255,255,255,.1);color:#ddd}
        .footer{text-align:center;padding:30px;color:#8296b0;font-size:14px}
        @media(max-width:1200px){.kpi-grid{grid-template-columns:repeat(3,1fr)}.chart-grid{grid-template-columns:1fr}}
        /* 数据质量提示样式 */
        .quality-issues{margin-top:15px;padding:15px;background:rgba(255,200,50,.1);border-radius:10px;color:#ffd54f;display:none}
        .quality-issues.show{display:block}
        .quality-issues li{list-style:none;padding:2px 0}
    </style>
</head>
<body>
<div class="dashboard">
    <div class="header">
        <div>
            <div class="title">西瑞集团经营数据驾驶舱</div>
            <div class="subtitle">金蝶EAS · 数据实时分析平台</div>
        </div>
        <div class="time" id="clock">--</div>
    </div>

    <div class="upload-panel">
        <div>
            <h3>销售数据导入</h3>
            <p style="color:#9fb5d1;margin-top:8px">支持金蝶EAS销售明细Excel自动分析</p>
        </div>
        <div>
            <span id="fileNameDisplay" class="file-name">未选择文件</span>
            <input type="file" id="fileInput" accept=".xlsx" style="display:none">
            <button class="upload-btn" onclick="document.getElementById('fileInput').click()">选择Excel</button>
            <button class="upload-btn" onclick="uploadFile()">上传数据</button>
        </div>
    </div>

    <!-- 数据质量提示 -->
    <div id="qualityIssues" class="quality-issues"></div>

    <div class="kpi-grid">
        <div class="kpi-card"><div class="kpi-label">今日销量</div><div class="kpi-value" id="k1">--</div><div class="unit">吨</div></div>
        <div class="kpi-card"><div class="kpi-label">本月销量</div><div class="kpi-value" id="k2">--</div><div class="unit">吨</div></div>
        <div class="kpi-card"><div class="kpi-label">本年销量</div><div class="kpi-value" id="k3">--</div><div class="unit">吨</div></div>
        <div class="kpi-card"><div class="kpi-label">今日销售额</div><div class="kpi-value" id="k4">--</div><div class="unit">万元</div></div>
        <div class="kpi-card"><div class="kpi-label">本月销售额</div><div class="kpi-value" id="k5">--</div><div class="unit">万元</div></div>
        <div class="kpi-card"><div class="kpi-label">本年销售额</div><div class="kpi-value" id="k6">--</div><div class="unit">万元</div></div>
    </div>

    <div class="chart-grid">
        <div class="chart-box"><div class="chart-title">销量趋势</div><div id="trendChart" class="chart"></div></div>
        <div class="chart-box"><div class="chart-title">销售额趋势</div><div id="amountChart" class="chart"></div></div>
    </div>
    <div class="chart-grid">
        <div class="chart-box"><div class="chart-title">月度销量分析</div><div id="monthChart" class="chart small-chart"></div></div>
        <div class="chart-box"><div class="chart-title">年度销量分析</div><div id="yearChart" class="chart small-chart"></div></div>
    </div>

    <div class="bottom-grid">
        <div class="rank-box"><div class="chart-title">产品销量TOP10</div><div id="productRank">暂无数据</div></div>
        <div class="rank-box"><div class="chart-title">客户销售TOP10</div><div id="customerRank">暂无数据</div></div>
    </div>

    <div class="footer">数据来源：金蝶EAS</div>
</div>

<script>
(function(){
    document.addEventListener('DOMContentLoaded', function(){
        // 时钟
        function updateClock(){
            var el=document.getElementById('clock');
            if(el) el.innerHTML=new Date().toLocaleString();
        }
        setInterval(updateClock,1000); updateClock();

        var fileInput=document.getElementById('fileInput');
        var fileNameDisplay=document.getElementById('fileNameDisplay');
        if(fileInput){
            fileInput.addEventListener('change', function(){
                if(this.files && this.files.length>0){
                    if(fileNameDisplay) fileNameDisplay.textContent=this.files[0].name;
                } else {
                    if(fileNameDisplay) fileNameDisplay.textContent='未选择文件';
                }
            });
        }

        window.uploadFile = async function(){
            var input=document.getElementById('fileInput');
            if(!input || !input.files || input.files.length===0){
                alert('请先选择Excel文件');
                return;
            }
            var file=input.files[0];
            var formData=new FormData();
            formData.append('file', file);
            try{
                var response=await fetch('/upload', {method:'POST', body:formData});
                var result=await response.json();
                if(result.success){
                    // 显示数据质量提示
                    var issuesDiv=document.getElementById('qualityIssues');
                    if(result.issues && result.issues.length>0){
                        var html='<strong>📊 数据质量检查：</strong><ul>';
                        for(var i=0;i<result.issues.length;i++){
                            html+='<li>'+result.issues[i]+'</li>';
                        }
                        html+='</ul>';
                        issuesDiv.innerHTML=html;
                        issuesDiv.className='quality-issues show';
                    } else {
                        issuesDiv.className='quality-issues';
                        issuesDiv.innerHTML='';
                    }
                    alert('✅ '+result.message);
                    loadAllData();
                } else {
                    alert('❌ 上传失败：'+result.message);
                }
            } catch(e){
                alert('❌ 网络错误：'+e.message);
            }
        };

        function formatNumber(num){ if(num===undefined || num===null) return '0'; return Number(num).toLocaleString(); }

        async function loadSummary(){
            try{
                var resp=await fetch('/api/summary');
                var d=await resp.json();
                var map={'k1':d.今日销量,'k2':d.本月销量,'k3':d.本年销量,'k4':d.今日销售额/10000,'k5':d.本月销售额/10000,'k6':d.本年销售额/10000};
                for(var id in map){ var el=document.getElementById(id); if(el) el.textContent=formatNumber(map[id]); }
            } catch(e){ console.error('加载KPI失败',e); }
        }

        function getBaseOption(){
            return {
                backgroundColor:'transparent',
                tooltip:{trigger:'axis', backgroundColor:'rgba(0,0,0,.75)', textStyle:{color:'#fff'}},
                grid:{top:60,bottom:70,left:90,right:50, containLabel:true},
                xAxis:{type:'category', axisLine:{lineStyle:{color:'#567'}}, axisLabel:{color:'#bdd', rotate:25}},
                yAxis:{type:'value', splitLine:{lineStyle:{color:'rgba(255,255,255,.1)'}}, axisLabel:{color:'#bdd'}}
            };
        }

        async function loadTrend(){
            try{
                var resp=await fetch('/api/data');
                var data=await resp.json();
                if(!data || data.length===0) return;
                var dates=data.map(function(item){return item.日期;});
                var sales=data.map(function(item){return item.总销量;});
                var chart=echarts.init(document.getElementById('trendChart'));
                var opt=getBaseOption();
                opt.xAxis.data=dates;
                opt.series=[{name:'销量',type:'line',smooth:true,symbol:'circle',symbolSize:8,data:sales,lineStyle:{width:4,color:'#d8a43a'},areaStyle:{color:'rgba(216,164,58,.25)'}}];
                chart.setOption(opt);
                window.addEventListener('resize', function(){ chart.resize(); });
            } catch(e){ console.error('加载趋势图失败',e); }
        }

        async function loadAmount(){
            try{
                var resp=await fetch('/api/data');
                var data=await resp.json();
                if(!data || data.length===0) return;
                var chart=echarts.init(document.getElementById('amountChart'));
                var opt=getBaseOption();
                opt.xAxis.data=data.map(function(item){return item.日期;});
                opt.series=[{name:'销售额(万元)',type:'bar',data:data.map(function(item){return item.总销售额/10000;}), itemStyle:{color:'#3fa66b', borderRadius:[6,6,0,0]}}];
                chart.setOption(opt);
                window.addEventListener('resize', function(){ chart.resize(); });
            } catch(e){ console.error('加载销售额趋势失败',e); }
        }

        async function loadMonth(){
            try{
                var resp=await fetch('/api/data');
                var data=await resp.json();
                if(!data || data.length===0) return;
                var monthMap={};
                data.forEach(function(item){ var m=item.日期.substring(0,7); monthMap[m]=(monthMap[m]||0)+item.总销量; });
                var chart=echarts.init(document.getElementById('monthChart'));
                var opt=getBaseOption();
                opt.xAxis.data=Object.keys(monthMap);
                opt.series=[{type:'bar',data:Object.values(monthMap),itemStyle:{color:'#4da3ff',borderRadius:[6,6,0,0]}}];
                chart.setOption(opt);
                window.addEventListener('resize', function(){ chart.resize(); });
            } catch(e){ console.error('加载月度分析失败',e); }
        }

        async function loadYear(){
            try{
                var resp=await fetch('/api/data');
                var data=await resp.json();
                if(!data || data.length===0) return;
                var yearMap={};
                data.forEach(function(item){ var y=item.日期.substring(0,4); yearMap[y]=(yearMap[y]||0)+item.总销量; });
                var chart=echarts.init(document.getElementById('yearChart'));
                var opt=getBaseOption();
                opt.xAxis.data=Object.keys(yearMap);
                opt.series=[{type:'bar',data:Object.values(yearMap),itemStyle:{color:'#f3c65a',borderRadius:[6,6,0,0]}}];
                chart.setOption(opt);
                window.addEventListener('resize', function(){ chart.resize(); });
            } catch(e){ console.error('加载年度分析失败',e); }
        }

        async function loadRank(){
            try{
                var productResp=await fetch('/api/product_rank');
                var productData=await productResp.json();
                var customerResp=await fetch('/api/customer_rank');
                var customerData=await customerResp.json();

                var productHTML='';
                if(productData && productData.length>0){
                    for(var i=0;i<productData.length;i++){
                        productHTML+='<div class="rank-item"><span>'+(i+1)+'. '+productData[i].name+'</span><b>'+formatNumber(productData[i].value)+' 吨</b></div>';
                    }
                } else { productHTML='暂无数据'; }
                document.getElementById('productRank').innerHTML=productHTML;

                var customerHTML='';
                if(customerData && customerData.length>0){
                    for(var j=0;j<customerData.length;j++){
                        customerHTML+='<div class="rank-item"><span>'+(j+1)+'. '+customerData[j].name+'</span><b>'+formatNumber(customerData[j].value/10000)+' 万元</b></div>';
                    }
                } else { customerHTML='暂无数据'; }
                document.getElementById('customerRank').innerHTML=customerHTML;
            } catch(e){ console.error('加载排行榜失败',e); }
        }

        window.loadAllData=function(){
            loadSummary();
            loadTrend();
            loadAmount();
            loadMonth();
            loadYear();
            loadRank();
        };

        loadAllData();
    });
})();
</script>
</body>
</html>
"""


# =====================================================
# 首页
# =====================================================

@app.get(
    "/",
    response_class=HTMLResponse
)
def home():
    return HTMLResponse(
        HTML_PAGE
    )


# =====================================================
# 启动
# =====================================================

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        app,
        host="0.0.0.0",
        port=8000
    )