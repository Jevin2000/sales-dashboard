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
    version="2.0"
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
# Excel字段智能识别
# =====================================================


def find_column(columns, keywords):

    for col in columns:

        text = str(col).lower()

        for k in keywords:

            if k.lower() in text:
                return col

    return None



# =====================================================
# 日期解析
# =====================================================


def parse_date(value):

    if pd.isna(value):
        return pd.NaT


    if isinstance(value, datetime):
        return value


    if isinstance(value, pd.Timestamp):
        return value



    value = str(value)



    formats = [

        "%Y-%m-%d",
        "%Y/%m/%d",
        "%Y%m%d",
        "%Y.%m.%d",
        "%d-%m-%Y"

    ]


    for fmt in formats:

        try:

            return pd.to_datetime(
                value,
                format=fmt
            )

        except:

            pass



    return pd.to_datetime(
        value,
        errors="coerce"
    )



# =====================================================
# Excel解析核心
# =====================================================


def parse_excel(df):


    columns = df.columns



    # 日期

    date_col = find_column(
        columns,
        [
            "日期",
            "时间",
            "date",
            "业务日期",
            "单据日期"
        ]
    )


    if not date_col:

        date_col = columns[0]



    # 数量

    qty_col = find_column(
        columns,
        [
            "数量",
            "销量",
            "出库数量",
            "qty",
            "吨"
        ]
    )


    if not qty_col:

        qty_col = columns[4]



    # 金额

    amt_col = find_column(
        columns,
        [
            "金额",
            "销售额",
            "价税合计",
            "amount",
            "含税"
        ]
    )


    if not amt_col:

        amt_col = columns[5]



    # 产品

    product_col = find_column(
        columns,
        [
            "产品",
            "物料",
            "商品",
            "品种"
        ]
    )


    # 客户

    customer_col = find_column(
        columns,
        [
            "客户",
            "购货单位",
            "单位",
            "往来单位"
        ]
    )



    result = pd.DataFrame()



    result["日期"] = (
        df[date_col]
        .apply(parse_date)
    )


    result["销量"] = pd.to_numeric(
        df[qty_col],
        errors="coerce"
    ).fillna(0)



    result["销售额"] = pd.to_numeric(
        df[amt_col],
        errors="coerce"
    ).fillna(0)



    if product_col:

        result["产品"] = (
            df[product_col]
            .astype(str)
        )

    else:

        result["产品"] = "未知产品"



    if customer_col:

        result["客户"] = (
            df[customer_col]
            .astype(str)
        )

    else:

        result["客户"] = "未知客户"




    result = result.dropna(
        subset=["日期"]
    )



    return result




# =====================================================
# 上传Excel
# =====================================================


@app.post("/upload")
async def upload_excel(
    file: UploadFile = File(...)
):

    try:


        content = await file.read()



        df = pd.read_excel(
            io.BytesIO(content)
        )



        data = parse_excel(df)



        # 保存原始数据

        raw_collection.delete_many({})

        raw_collection.insert_many(
            data.to_dict("records")
        )



        # 日汇总

        daily = (

            data
            .groupby(
                data["日期"].dt.date
            )
            .agg(
                {
                    "销量":"sum",
                    "销售额":"sum"
                }
            )
            .reset_index()

        )


        daily.columns = [
            "日期",
            "总销量",
            "总销售额"
        ]



        collection.delete_many({})



        records = (
            daily
            .to_dict("records")
        )


        for r in records:

            r["日期"] = (
                str(r["日期"])
            )



        collection.insert_many(
            records
        )



        return {

            "success":True,

            "message":
            f"成功导入 {len(records)} 天数据"

        }



    except Exception as e:


        return JSONResponse(

            {
                "success":False,
                "message":str(e)
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
            {"_id":0}
        )
    )


    return data



# =====================================================
# KPI接口
# =====================================================


@app.get("/api/summary")
def summary():


    data=list(
        collection.find(
            {},
            {"_id":0}
        )
    )


    if not data:

        return {}


    df=pd.DataFrame(data)


    df["日期"]=pd.to_datetime(
        df["日期"]
    )


    now=pd.Timestamp.now()


    return {


        "今日销量":

        int(
            df[
                df["日期"].dt.date
                ==
                now.date()
            ]["总销量"].sum()
        ),


        "本月销量":

        int(
            df[
                df["日期"].dt.month
                ==
                now.month
            ]["总销量"].sum()
        ),


        "本年销量":

        int(
            df[
                df["日期"].dt.year
                ==
                now.year
            ]["总销量"].sum()
        ),



        "今日销售额":

        float(
            df[
                df["日期"].dt.date
                ==
                now.date()
            ]["总销售额"].sum()
        ),


        "本月销售额":

        float(
            df[
                df["日期"].dt.month
                ==
                now.month
            ]["总销售额"].sum()
        ),


        "本年销售额":

        float(
            df[
                df["日期"].dt.year
                ==
                now.year
            ]["总销售额"].sum()
        )

    }



# =====================================================
# 产品排行
# =====================================================


@app.get("/api/product_rank")
def product_rank():


    data=list(
        raw_collection.find(
            {},
            {"_id":0}
        )
    )


    if not data:

        return []



    df=pd.DataFrame(data)


    result=(

        df.groupby("产品")
        ["销量"]
        .sum()
        .sort_values(
            ascending=False
        )
        .head(10)

    )


    return [

        {
            "name":k,
            "value":float(v)
        }

        for k,v in result.items()

    ]



# =====================================================
# 客户排行
# =====================================================


@app.get("/api/customer_rank")
def customer_rank():


    data=list(
        raw_collection.find(
            {},
            {"_id":0}
        )
    )


    if not data:

        return []



    df=pd.DataFrame(data)



    result=(

        df.groupby("客户")
        ["销售额"]
        .sum()
        .sort_values(
            ascending=False
        )
        .head(10)

    )



    return [

        {
            "name":k,
            "value":float(v)
        }

        for k,v in result.items()

    ]


# =====================================================
# 企业驾驶舱HTML（完整修复版）
# =====================================================

HTML_PAGE = r"""

<!DOCTYPE html>

<html lang="zh-CN">


<head>

<meta charset="UTF-8">

<meta name="viewport"
content="width=device-width,initial-scale=1.0">


<title>
西瑞集团经营数据驾驶舱
</title>


<script src="https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js">
</script>



<style>


/* =============================
   基础
============================= */


*{

margin:0;
padding:0;
box-sizing:border-box;

}


body{


background:

linear-gradient(
135deg,
#071426,
#0d2340
);


font-family:

"Microsoft YaHei",
Arial,
sans-serif;


color:white;


min-height:100vh;


}




.dashboard{


width:95%;

max-width:1920px;

margin:auto;


padding:25px;


}




/* =============================
顶部
============================= */


.header{


height:100px;


display:flex;

justify-content:space-between;

align-items:center;


border-bottom:

1px solid rgba(255,255,255,.15);


}



.title{


font-size:38px;

font-weight:bold;


letter-spacing:6px;


color:#f3c65a;


}



.subtitle{


font-size:16px;

color:#9fb5d1;


margin-top:10px;


}




.time{


font-size:18px;

color:#9fd5ff;


}





/* =============================
上传区域
============================= */


.upload-panel{


margin-top:25px;


background:

rgba(255,255,255,.06);


border-radius:15px;


padding:20px;


display:flex;

justify-content:space-between;

align-items:center;


}



.upload-btn{


background:#d8a43a;

border:none;


padding:12px 35px;


border-radius:8px;


font-size:16px;


cursor:pointer;


font-weight:bold;


}



.upload-btn:hover{


opacity:.85;


}
.file-name{


margin-right:20px;

color:#9fd5ff;


}




/* =============================
KPI
============================= */


.kpi-grid{


margin-top:25px;


display:grid;


grid-template-columns:

repeat(6,1fr);


gap:20px;


}



.kpi-card{


background:

linear-gradient(
145deg,
rgba(255,255,255,.10),
rgba(255,255,255,.03)
);


border-radius:16px;


padding:25px;


text-align:center;


border:

1px solid rgba(255,255,255,.12);


}



.kpi-label{


font-size:16px;

color:#9fb5d1;


}



.kpi-value{


margin-top:15px;


font-size:38px;


font-weight:bold;


color:#ffffff;


}



.unit{


font-size:15px;


color:#d8a43a;


}




/* =============================
图表
============================= */


.chart-grid{


margin-top:25px;


display:grid;


grid-template-columns:

2fr 1fr;


gap:25px;


}



.chart-box{


background:

rgba(255,255,255,.06);


border-radius:15px;


padding:20px;


border:

1px solid rgba(255,255,255,.12);


}




.chart-title{


font-size:18px;


color:#f3c65a;


margin-bottom:15px;


}



.chart{


height:380px;


width:100%;


}





.small-chart{


height:300px;


}

/* =============================
底部分析
============================= */


.bottom-grid{


margin-top:25px;


display:grid;


grid-template-columns:

1fr 1fr;


gap:25px;


}



.rank-box{


background:

rgba(255,255,255,.06);


border-radius:15px;


padding:20px;


}



.rank-item{


display:flex;


justify-content:space-between;


padding:12px 0;


border-bottom:

1px solid rgba(255,255,255,.1);


color:#ddd;


}




.footer{


text-align:center;


padding:30px;


color:#8296b0;


font-size:14px;


}




@media(max-width:1200px){


.kpi-grid{

grid-template-columns:

repeat(3,1fr);

}



.chart-grid{

grid-template-columns:1fr;

}




}


</style>


</head>




<body>


<div class="dashboard">



<!-- 顶部 -->

<div class="header">


<div>


<div class="title">

西瑞集团经营数据驾驶舱

</div>


<div class="subtitle">

金蝶EAS · 数据实时分析平台

</div>


</div>



<div class="time" id="clock">

--

</div>


</div>





<!-- 上传 -->


<div class="upload-panel">


<div>


<h3>

销售数据导入

</h3>


<p style="color:#9fb5d1;margin-top:8px">

支持金蝶EAS销售明细Excel自动分析

</p>


</div>



<div>


<span id="fileNameDisplay"
class="file-name">

未选择文件

</span>



<input

type="file"

id="fileInput"

accept=".xlsx"

style="display:none">



<button

class="upload-btn"

onclick="document.getElementById('fileInput').click()">

选择Excel

</button>



<button

class="upload-btn"

onclick="uploadFile()">

上传数据

</button>



</div>


</div>






<!-- KPI -->


<div class="kpi-grid">


<div class="kpi-card">

<div class="kpi-label">

今日销量

</div>

<div class="kpi-value" id="k1">

--

</div>

<div class="unit">

吨

</div>

</div>




<div class="kpi-card">

<div class="kpi-label">

本月销量

</div>

<div class="kpi-value" id="k2">

--

</div>

<div class="unit">

吨

</div>

</div>




<div class="kpi-card">

<div class="kpi-label">

本年销量

</div>

<div class="kpi-value" id="k3">

--

</div>

<div class="unit">

吨

</div>

</div>




<div class="kpi-card">

<div class="kpi-label">

今日销售额

</div>

<div class="kpi-value" id="k4">

--

</div>

<div class="unit">

万元

</div>

</div>




<div class="kpi-card">

<div class="kpi-label">

本月销售额

</div>

<div class="kpi-value" id="k5">

--

</div>

<div class="unit">

万元

</div>

</div>




<div class="kpi-card">

<div class="kpi-label">

本年销售额

</div>

<div class="kpi-value" id="k6">

--

</div>

<div class="unit">

万元

</div>

</div>


</div>





<!-- 图表 -->



<div class="chart-grid">



<div class="chart-box">


<div class="chart-title">

年度销量趋势

</div>


<div id="trendChart"
class="chart">

</div>


</div>




<div class="chart-box">


<div class="chart-title">

销售额趋势

</div>


<div id="amountChart"
class="chart">

</div>


</div>



</div>






<div class="chart-grid">


<div class="chart-box">


<div class="chart-title">

月度销量分析

</div>


<div id="monthChart"
class="chart small-chart">

</div>


</div>




<div class="chart-box">


<div class="chart-title">

年度销量分析

</div>


<div id="yearChart"
class="chart small-chart">

</div>


</div>



</div>






<div class="bottom-grid">



<div class="rank-box">


<div class="chart-title">

产品销量TOP10

</div>


<div id="productRank">

暂无数据

</div>


</div>




<div class="rank-box">


<div class="chart-title">

客户销售TOP10

</div>


<div id="customerRank">

暂无数据

</div>


</div>




</div>





<div class="footer">


数据来源：金蝶EAS


</div>



</div>



<script>
// ============================================================
// 完全修复版 JavaScript
// ============================================================

(function() {
    "use strict";

    // ---------- 等待DOM加载 ----------
    document.addEventListener('DOMContentLoaded', function() {

        // ---------- 时钟 ----------
        function updateClock() {
            var el = document.getElementById('clock');
            if (el) {
                el.innerHTML = new Date().toLocaleString();
            }
        }
        setInterval(updateClock, 1000);
        updateClock();

        // ---------- 文件选择显示 ----------
        var fileInput = document.getElementById('fileInput');
        var fileNameDisplay = document.getElementById('fileNameDisplay');

        if (fileInput) {
            fileInput.addEventListener('change', function() {
                if (this.files && this.files.length > 0) {
                    if (fileNameDisplay) {
                        fileNameDisplay.textContent = this.files[0].name;
                    }
                } else {
                    if (fileNameDisplay) {
                        fileNameDisplay.textContent = '未选择文件';
                    }
                }
            });
        }

        // ---------- 上传函数（暴露到全局） ----------
        window.uploadFile = async function() {
            var input = document.getElementById('fileInput');
            if (!input || !input.files || input.files.length === 0) {
                alert('请先选择Excel文件');
                return;
            }

            var file = input.files[0];
            var formData = new FormData();
            formData.append('file', file);

            try {
                var response = await fetch('/upload', {
                    method: 'POST',
                    body: formData
                });
                var result = await response.json();
                if (result.success) {
                    alert('✅ ' + result.message);
                    loadAllData();
                } else {
                    alert('❌ 上传失败：' + result.message);
                }
            } catch (e) {
                alert('❌ 网络错误：' + e.message);
            }
        };

        // ---------- 工具函数 ----------
        function formatNumber(num) {
            if (num === undefined || num === null) return '0';
            return Number(num).toLocaleString();
        }

        // ---------- 加载KPI ----------
        async function loadSummary() {
            try {
                var resp = await fetch('/api/summary');
                var d = await resp.json();
                var map = {
                    'k1': d.今日销量,
                    'k2': d.本月销量,
                    'k3': d.本年销量,
                    'k4': d.今日销售额 / 10000,
                    'k5': d.本月销售额 / 10000,
                    'k6': d.本年销售额 / 10000
                };
                for (var id in map) {
                    var el = document.getElementById(id);
                    if (el) {
                        el.textContent = formatNumber(map[id]);
                    }
                }
            } catch (e) {
                console.error('加载KPI失败', e);
            }
        }

        // ---------- 图表基础配置 ----------
        function getBaseOption() {
            return {
                backgroundColor: 'transparent',
                tooltip: {
                    trigger: 'axis',
                    backgroundColor: 'rgba(0,0,0,.75)',
                    textStyle: { color: '#fff' }
                },
                grid: {
                    top: 60,
                    bottom: 70,
                    left: 90,
                    right: 50,
                    containLabel: true
                },
                xAxis: {
                    type: 'category',
                    axisLine: { lineStyle: { color: '#567' } },
                    axisLabel: { color: '#bdd', rotate: 25 }
                },
                yAxis: {
                    type: 'value',
                    splitLine: { lineStyle: { color: 'rgba(255,255,255,.1)' } },
                    axisLabel: { color: '#bdd' }
                }
            };
        }

        // ---------- 加载趋势图 ----------
        async function loadTrend() {
            try {
                var resp = await fetch('/api/data');
                var data = await resp.json();
                if (!data || data.length === 0) return;

                var dates = data.map(function(item) { return item.日期; });
                var sales = data.map(function(item) { return item.总销量; });

                var chart = echarts.init(document.getElementById('trendChart'));
                var opt = getBaseOption();
                opt.xAxis.data = dates;
                opt.series = [{
                    name: '销量',
                    type: 'line',
                    smooth: true,
                    symbol: 'circle',
                    symbolSize: 8,
                    data: sales,
                    lineStyle: { width: 4, color: '#d8a43a' },
                    areaStyle: { color: 'rgba(216,164,58,.25)' }
                }];
                chart.setOption(opt);
                window.addEventListener('resize', function() { chart.resize(); });
            } catch (e) {
                console.error('加载趋势图失败', e);
            }
        }

        // ---------- 加载销售额趋势 ----------
        async function loadAmount() {
            try {
                var resp = await fetch('/api/data');
                var data = await resp.json();
                if (!data || data.length === 0) return;

                var chart = echarts.init(document.getElementById('amountChart'));
                var opt = getBaseOption();
                opt.xAxis.data = data.map(function(item) { return item.日期; });
                opt.series = [{
                    name: '销售额(万元)',
                    type: 'bar',
                    data: data.map(function(item) { return item.总销售额 / 10000; }),
                    itemStyle: { color: '#3fa66b', borderRadius: [6, 6, 0, 0] }
                }];
                chart.setOption(opt);
                window.addEventListener('resize', function() { chart.resize(); });
            } catch (e) {
                console.error('加载销售额趋势失败', e);
            }
        }

        // ---------- 加载月度分析 ----------
        async function loadMonth() {
            try {
                var resp = await fetch('/api/data');
                var data = await resp.json();
                if (!data || data.length === 0) return;

                var monthMap = {};
                data.forEach(function(item) {
                    var m = item.日期.substring(0, 7);
                    monthMap[m] = (monthMap[m] || 0) + item.总销量;
                });

                var chart = echarts.init(document.getElementById('monthChart'));
                var opt = getBaseOption();
                opt.xAxis.data = Object.keys(monthMap);
                opt.series = [{
                    type: 'bar',
                    data: Object.values(monthMap),
                    itemStyle: { color: '#4da3ff', borderRadius: [6, 6, 0, 0] }
                }];
                chart.setOption(opt);
                window.addEventListener('resize', function() { chart.resize(); });
            } catch (e) {
                console.error('加载月度分析失败', e);
            }
        }

        // ---------- 加载年度分析 ----------
        async function loadYear() {
            try {
                var resp = await fetch('/api/data');
                var data = await resp.json();
                if (!data || data.length === 0) return;

                var yearMap = {};
                data.forEach(function(item) {
                    var y = item.日期.substring(0, 4);
                    yearMap[y] = (yearMap[y] || 0) + item.总销量;
                });

                var chart = echarts.init(document.getElementById('yearChart'));
                var opt = getBaseOption();
                opt.xAxis.data = Object.keys(yearMap);
                opt.series = [{
                    type: 'bar',
                    data: Object.values(yearMap),
                    itemStyle: { color: '#f3c65a', borderRadius: [6, 6, 0, 0] }
                }];
                chart.setOption(opt);
                window.addEventListener('resize', function() { chart.resize(); });
            } catch (e) {
                console.error('加载年度分析失败', e);
            }
        }

        // ---------- 加载排行榜 ----------
        async function loadRank() {
            try {
                var productResp = await fetch('/api/product_rank');
                var productData = await productResp.json();
                var customerResp = await fetch('/api/customer_rank');
                var customerData = await customerResp.json();

                var productHTML = '';
                if (productData && productData.length > 0) {
                    for (var i = 0; i < productData.length; i++) {
                        productHTML += '<div class="rank-item">' +
                            '<span>' + (i + 1) + '. ' + productData[i].name + '</span>' +
                            '<b>' + formatNumber(productData[i].value) + ' 吨</b>' +
                            '</div>';
                    }
                } else {
                    productHTML = '暂无数据';
                }
                document.getElementById('productRank').innerHTML = productHTML;

                var customerHTML = '';
                if (customerData && customerData.length > 0) {
                    for (var j = 0; j < customerData.length; j++) {
                        customerHTML += '<div class="rank-item">' +
                            '<span>' + (j + 1) + '. ' + customerData[j].name + '</span>' +
                            '<b>' + formatNumber(customerData[j].value / 10000) + ' 万元</b>' +
                            '</div>';
                    }
                } else {
                    customerHTML = '暂无数据';
                }
                document.getElementById('customerRank').innerHTML = customerHTML;
            } catch (e) {
                console.error('加载排行榜失败', e);
            }
        }

        // ---------- 统一加载函数 ----------
        window.loadAllData = function() {
            loadSummary();
            loadTrend();
            loadAmount();
            loadMonth();
            loadYear();
            loadRank();
        };

        // ---------- 启动 ----------
        loadAllData();

    }); // DOMContentLoaded
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


if __name__=="__main__":


    import uvicorn


    uvicorn.run(

        app,

        host="0.0.0.0",

        port=8000

    )