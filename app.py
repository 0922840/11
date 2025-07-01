streamlit
pandas
prophet
matplotlib
openpyxl
import streamlit as st
import pandas as pd
from prophet import Prophet
import matplotlib.pyplot as plt
from io import BytesIO
import matplotlib
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

# 中文显示设置
matplotlib.rcParams['font.sans-serif'] = ['SimHei']
matplotlib.rcParams['axes.unicode_minus'] = False

# 页面设置
st.set_page_config(page_title="订单预测系统", layout="wide", page_icon="📦")

# 顶部标题
st.title("📦 订单预测与高峰调度系统（增强美观版）")
st.markdown("更智能的物流预测系统，助你优化排班与资源配置")

# 邮件配置区域（侧边栏）
with st.sidebar:
    st.header("📧 邮件配置（可选）")
    smtp_server = st.text_input("SMTP服务器", value="smtp.qq.com")
    smtp_port = st.number_input("SMTP端口", value=465, step=1)
    sender_email = st.text_input("发件邮箱")
    sender_password = st.text_input("发件邮箱密码（授权码）", type="password")
    admin_email = st.text_input("接收邮箱", value="2103225537@qq.com")

    st.header("⚙️ 参数配置")
    num_workers = st.number_input("配布员人数", value=18, step=1)
    efficiency_per_hour = st.number_input("每人每小时效率（卷）", value=80, step=10)
    peak_hours = st.selectbox("高峰时段（小时）", [1, 2, 3], index=1)
    sku_efficiency = st.slider("SKU效率系数", 0.5, 1.0, 0.83, step=0.01)
    safety_factor = st.slider("安全系数", 1.0, 1.5, 1.1, step=0.01)
    peak_coef = st.slider("峰值系数", 1.0, 2.5, 1.82, step=0.01)
    use_peak = st.checkbox("启用淡旺季识别", value=True)

# 上传数据
st.markdown("### 📤 上传订单数据（Excel格式）")
uploaded_file = st.file_uploader("文件需包含“日期”和“出库量”列", type=["xlsx"])

# 邮件发送函数
def send_email_with_attachments(smtp_server, smtp_port, sender_email, sender_password,
                                receiver_email, subject, body_text, attachments: dict):
    try:
        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = receiver_email
        msg["Subject"] = subject
        msg.attach(MIMEText(body_text, "plain", "utf-8"))

        for fname, fcontent in attachments.items():
            if fname.endswith(".png"):
                img = MIMEImage(fcontent)
                img.add_header('Content-Disposition', 'attachment', filename=fname)
                msg.attach(img)
            else:
                part = MIMEApplication(fcontent)
                part.add_header('Content-Disposition', 'attachment', filename=fname)
                msg.attach(part)

        with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
            server.set_debuglevel(1)
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, receiver_email, msg.as_string())
        return True, "✅ 邮件发送成功"
    except Exception as e:
        return False, f"⚠️ 邮件发送失败：{str(e)}"

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    if '日期' in df.columns and '出库量' in df.columns:
        df = df.rename(columns={'日期': 'ds', '出库量': 'y'})
        df['ds'] = pd.to_datetime(df['ds'])

        if use_peak:
            df['is_peak'] = df['ds'].dt.month.isin([3, 4, 10, 11]).astype(int)

        model = Prophet()
        if use_peak:
            model.add_regressor('is_peak')
        model.fit(df)

        future = model.make_future_dataframe(periods=7)
        if use_peak:
            future['is_peak'] = future['ds'].dt.month.isin([3, 4, 10, 11]).astype(int)
        forecast = model.predict(future)[['ds', 'yhat']]
        forecast['yhat'] = forecast['yhat'].round(0).astype(int)
        forecast = forecast[forecast['ds'] > df['ds'].max()]
        forecast = forecast.rename(columns={'yhat': '预测订单量（卷）'})

        limit = int(num_workers * efficiency_per_hour * peak_hours * sku_efficiency * safety_factor)
        forecast['高峰负荷（卷）'] = (forecast['预测订单量（卷）'] * peak_coef / 3).round(0).astype(int)
        forecast['高峰处理上限（卷）'] = limit
        forecast['是否触发错峰策略'] = forecast['高峰负荷（卷）'] > limit
        forecast['是否触发错峰策略'] = forecast['是否触发错峰策略'].replace({True: '是', False: '否'})
        forecast['人力资源利用率'] = (forecast['预测订单量（卷）'] / (num_workers * 8 * efficiency_per_hour * sku_efficiency)).round(2)
        forecast['配送车辆利用率'] = (forecast['预测订单量（卷）'] / 9000).clip(upper=1.0).round(2)
        forecast['叉车资源利用率'] = (forecast['预测订单量（卷）'] / 10800).clip(upper=1.0).round(2)

        # 推荐策略计算
        def calculate_recommendations(row):
            load = row['高峰负荷（卷）']
            if load <= limit:
                return 0, 0, "否", "正常时段"
            diff = load - limit
            add_workers = int(diff / (efficiency_per_hour * sku_efficiency * peak_hours)) + 1
            extended_hours = round(diff / (num_workers * efficiency_per_hour * sku_efficiency), 2)
            batch_flag = "是" if add_workers > num_workers * 0.3 else "否"
            rec_time = "非高峰时段" if batch_flag == "是" else "延长作业时间"
            return add_workers, extended_hours, batch_flag, rec_time

        forecast[['建议增派人数', '建议延长小时数', '是否分批', '推荐发货时间段']] = forecast.apply(
            lambda row: pd.Series(calculate_recommendations(row)), axis=1
        )

        st.markdown(f"#### 🎯 当前高峰处理上限：**{limit} 卷**")
        st.dataframe(forecast.rename(columns={'ds': '日期'}).reset_index(drop=True), use_container_width=True)

        with st.expander("📋 智能策略建议（点击查看）", expanded=True):
            for _, row in forecast.iterrows():
                date = row['ds'].date()
                if row['是否触发错峰策略'] == '是':
                    st.warning(f"📌 {date}：高峰负荷 {row['高峰负荷（卷）']} 卷，建议增派 **{row['建议增派人数']}人**，或延长作业 **{row['建议延长小时数']}小时**，是否分批：{row['是否分批']}，推荐时间段：{row['推荐发货时间段']}")
                else:
                    st.success(f"✅ {date}：高峰负荷 {row['高峰负荷（卷）']} 卷，系统可控")

        # 图表展示
        st.markdown("### 📊 订单趋势图")
        fig, ax = plt.subplots(figsize=(12, 5))
        ax.plot(df['ds'], df['y'], label='历史订单', marker='o')
        ax.plot(forecast['ds'], forecast['预测订单量（卷）'], label='预测订单', linestyle='--', marker='o')
        ax.plot(forecast['ds'], forecast['高峰负荷（卷）'], label='高峰负荷', linestyle='--', color='red')
        ax.axhline(y=limit, color='green', linestyle='-', label='高峰处理上限')
        ax.set_xlabel("日期")
        ax.set_ylabel("订单量（卷）")
        ax.set_title("📈 订单趋势图")
        ax.legend()
        plt.xticks(rotation=45)
        plt.tight_layout()
        st.pyplot(fig)

        # 邮件附件准备
        buf = BytesIO()
        fig.savefig(buf, format="png")
        buf.seek(0)
        png_content = buf.getvalue()

        excel_buf = BytesIO()
        with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer:
            forecast.to_excel(writer, index=False)
        excel_buf.seek(0)
        excel_content = excel_buf.getvalue()

        def get_text_advice(forecast_df):
            lines = ["订单预测与建议："]
            for _, row in forecast_df.iterrows():
                flag_mark = '×' if row['是否触发错峰策略'] == '是' else '✔'
                lines.append(
                    f"{row['ds'].date()}：预测订单：{row['预测订单量（卷）']}，负荷：{row['高峰负荷（卷）']}，"
                    f"策略触发：{row['是否触发错峰策略']} {flag_mark}，增员：{row['建议增派人数']}，延长：{row['建议延长小时数']} 小时，建议：{row['推荐发货时间段']}"
                )
            return "\n".join(lines)

        st.markdown("### 📤 发送预测邮件")
        if st.button("📨 发送邮件"):
            if all([smtp_server, sender_email, sender_password, admin_email]):
                advice_text = get_text_advice(forecast)
                success, msg = send_email_with_attachments(
                    smtp_server, smtp_port, sender_email, sender_password,
                    admin_email, "订单预测与调度建议",
                    advice_text,
                    attachments={
                        "预测结果图.png": png_content,
                        "预测明细.xlsx": excel_content
                    }
                )
                st.success(msg) if success else st.error(msg)
            else:
                st.error("请填写完整的邮件配置信息！")
    else:
        st.error("Excel文件中缺少 '日期' 或 '出库量' 列，请检查格式")
else:
    st.info("请上传含“日期”和“出库量”的Excel文件用于预测")
