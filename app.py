import streamlit as st
import pandas as pd
import pymongo
from datetime import datetime
import plotly.express as px
import os

# 1. 连接MongoDB（从Render环境变量读取密码）
client = pymongo.MongoClient(os.getenv("MONGO_URI"))
db = client["sales_db"]  # 这里的数据会存放在你现有集群的“sales_db”库里，不会影响你之前的项目
collection = db["daily_data"]

st.set_page_config(layout="wide")
st.title("📊 企业产销量数据看板")

# ---------- 侧边栏：上传Excel ----------
with st.sidebar:
    st.header("📤 上传数据")
    st.markdown("请确保Excel包含：**日期、产品、销量、产量** 四列")
    uploaded_file = st.file_uploader("选择Excel文件（.xlsx）", type=["xlsx"])
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        required_cols = ["日期", "产品", "销量", "产量"]
        if all(col in df.columns for col in required_cols):
            records = df.to_dict("records")
            for rec in records:
                rec["日期"] = pd.to_datetime(rec["日期"]).strftime("%Y-%m-%d")
            collection.insert_many(records)
            st.success(f"✅ 成功上传 {len(records)} 条数据！")
        else:
            st.error(f"❌ Excel必须包含列：{', '.join(required_cols)}")

# ---------- 读取数据 ----------
data = list(collection.find({}, {"_id": 0}))
if not data:
    st.info("💡 暂无数据，请先在左侧上传Excel文件")
    st.stop()

df_all = pd.DataFrame(data)
df_all["日期"] = pd.to_datetime(df_all["日期"])

# 汇总日、月、年
df_daily = df_all.groupby("日期").agg({"销量": "sum", "产量": "sum"}).reset_index()
df_daily["月份"] = df_daily["日期"].dt.to_period("M")
df_monthly = df_daily.groupby("月份").agg({"销量": "sum", "产量": "sum"}).reset_index()
df_monthly["月份"] = df_monthly["月份"].astype(str)
df_yearly = df_monthly.groupby(df_monthly["月份"].str[:4]).agg({"销量": "sum", "产量": "sum"}).reset_index()
df_yearly.columns = ["年份", "销量", "产量"]

# ---------- 指标卡片 ----------
col1, col2, col3 = st.columns(3)
col1.metric("📅 今日销量", f"{df_daily[df_daily['日期'] == df_daily['日期'].max()]['销量'].sum():,.0f}")
col2.metric("📆 本月销量", f"{df_monthly[df_monthly['月份'] == df_monthly['月份'].max()]['销量'].sum():,.0f}")
col3.metric("📈 本年销量", f"{df_yearly[df_yearly['年份'] == df_yearly['年份'].max()]['销量'].sum():,.0f}")

# ---------- 图表 ----------
st.subheader("📈 日销量与产量趋势")
fig1 = px.line(df_daily, x="日期", y=["销量", "产量"], markers=True)
st.plotly_chart(fig1, use_container_width=True)

st.subheader("📊 月度销量与产量")
fig2 = px.bar(df_monthly, x="月份", y=["销量", "产量"], barmode="group")
st.plotly_chart(fig2, use_container_width=True)

st.subheader("📊 年度销量与产量")
fig3 = px.bar(df_yearly, x="年份", y=["销量", "产量"], barmode="group")
st.plotly_chart(fig3, use_container_width=True)

st.subheader("📋 原始数据明细")
st.dataframe(df_all, use_container_width=True, height=300)