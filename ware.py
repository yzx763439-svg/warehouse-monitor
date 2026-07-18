import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime
import io

# --- 1. 初始化数据库 (保证数据永久保存、多端动态读取) ---
def init_db():
    conn = sqlite3.connect('warehouse.db')
    cursor = conn.cursor()
    # 创建采购主数据表
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS po_orders (
            po_id TEXT PRIMARY KEY,
            vendor TEXT,
            tracking_no TEXT,
            sku TEXT,
            expected_qty INTEGER,
            actual_qty INTEGER,
            status TEXT,
            arrival_time TEXT,
            is_urgent TEXT,
            notes TEXT
        )
    ''')
    # 创建操作日志表
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS action_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT,
            log_type TEXT,
            message TEXT
        )
    ''')
    conn.commit()
    conn.close()

init_db()

# --- 2. 数据库操作核心工具函数 ---
def get_db_connection():
    return sqlite3.connect('warehouse.db')

def log_action(log_type, message):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("INSERT INTO action_logs (timestamp, log_type, message) VALUES (?, ?, ?)",
                   (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), log_type, message))
    conn.commit()
    conn.close()

# --- 3. 页面全局配置 ---
st.set_page_config(page_title="动态仓库监控平台", layout="wide")

# 侧边栏：切换不同的工作岗位（多端协同模拟）
st.sidebar.title("🏢 仓储多端协同系统")
role = st.sidebar.radio("请选择当前操作岗位：", ["📊 实时监控大屏", "👩‍💻 采购部（Excel导入）", "📦 收货台（扫码作业）"])

# --- 岗位一：实时监控大屏（全动态自动刷新） ---
if role == "📊 实时监控大屏":
    st.title("🏭 仓库各流程实时数据监控大屏")
    
    # 💡 动态核心：利用 Streamlit 的 fragment 或 rerun 机制实现自动刷新
    # 这里设置每 5 秒，大屏自动重新读取数据库，实现真正动态
    st.caption("🔄 大屏正在实时监控中... 每 5 秒自动刷新数据")
    
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT * FROM po_orders", conn)
    df_logs = pd.read_sql_query("SELECT * FROM action_logs ORDER BY id DESC LIMIT 5", conn)
    conn.close()

    if df.empty:
        st.info("💡 暂无动态数据。请先前往【采购部】页面导入采购单 Excel。")
    else:
        # 1. 动态看板数字统计
        total_po = len(df)
        wait_receive = len(df[df["status"] == "待收货"])
        received = len(df[df["status"] == "已到货"])
        urgent_count = len(df[(df["is_urgent"] == "是") & (df["status"] == "待收货")])

        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric(label="📋 总采购单量", value=f"{total_po} 单")
        kpi2.metric(label="🚚 待收货单数", value=f"{wait_receive} 单")
        kpi3.metric(label="✅ 今日已到货", value=f"{received} 单")
        kpi4.metric(label="🚨 待处理加急单", value=f"{urgent_count} 单", delta="优先处理" if urgent_count > 0 else None, delta_color="inverse")

        st.markdown("---")

        # 2. 左右分栏：实时流水与图表
        col_chart, col_log = st.columns([1, 1])
        with col_chart:
            st.subheader("⏱️ 环节时效监控")
            # 动态生成当前状态对比图
            status_counts = df["status"].value_counts()
            st.bar_chart(status_counts)
        
        with col_log:
            st.subheader("🔔 实时操作动态流水")
            for _, log in df_logs.iterrows():
                if "🚨" in log['message']:
                    st.error(f"[{log['timestamp']}] {log['message']}")
                else:
                    st.code(f"[{log['timestamp']}] {log['message']}")

        st.markdown("---")
        # 3. 动态全链路状态监控表
        st.subheader("📋 采购单全链路实时状态监控表")
        
        def style_rows(row):
            if row["is_urgent"] == "是" and row["status"] == "待收货":
                return ['background-color: #ffcccc; color: #cc0000; font-weight: bold'] * len(row)
            elif row["status"] == "已到货":
                return ['background-color: #e6f4ea; color: #137333'] * len(row)
            return [''] * len(row)
        
        st.dataframe(df.style.apply(style_rows, axis=1), use_container_width=True)

    # 5秒后自动重跑代码，刷新大屏数据
    time.sleep(5)
    st.rerun()

# --- 岗位二：采购部（真正支持 Excel 批量导入及校验） ---
elif role == "👩‍💻 采购部（Excel导入）":
    st.title("👩‍💻 采购主数据源导入中心")
    st.write("根据操作数字化需求，支持采购单（PO号、供应商、物流单号、SKU、数量等）批量导入及格式校验。")

    # 1. 提供一个动态下载模板的功能（方便测试）
    st.subheader("1. 下载标准导入模板")
    template_df = pd.DataFrame([
        {"PO号": "PO20260001", "供应商": "某某工厂", "物流单号": "SF123456789", "SKU": "SKU-A", "预计数量": 100, "是否加急": "是"},
        {"PO号": "PO20260002", "供应商": "某某贸易商", "物流单号": "YT987654321", "SKU": "SKU-B", "预计数量": 50, "是否加急": "否"}
    ])
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        template_df.to_excel(writer, index=False, sheet_name='Sheet1')
    st.download_button(label="📥 点击下载 Excel 导入模板.xlsx", data=buffer.getvalue(), file_name="采购单导入模板.xlsx", mime="application/vnd.ms-excel")

    st.markdown("---")

    # 2. 上传并解析 Excel 文件
    st.subheader("2. 上传采购单 Excel 进行校验导入")
    uploaded_file = st.file_uploader("选择写好数据的采购单 Excel 文件", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        try:
            input_df = pd.read_excel(uploaded_file)
            st.write("📋 识别到的上传数据如下：", input_df)
            
            if st.button("🚀 确认校验并导入数据库"):
                conn = get_db_connection()
                cursor = conn.cursor()
                
                success_count = 0
                skip_count = 0
                
                for _, row in input_df.iterrows():
                    po_id = str(row["PO号"]).strip()
                    # 校验重复PO号
                    cursor.execute("SELECT po_id FROM po_orders WHERE po_id = ?", (po_id,))
                    exists = cursor.fetchone()
                    
                    if exists:
                        skip_count += 1
                        continue # 重复则跳过（可根据需求改为覆盖）
                    
                    cursor.execute('''
                        INSERT INTO po_orders (po_id, vendor, tracking_no, sku, expected_qty, actual_qty, status, arrival_time, is_urgent, notes)
                        VALUES (?, ?, ?, ?, ?, 0, '待收货', '-', ?, '-')
                    ''', (po_id, str(row["供应商"]), str(row["物流单号"]), str(row["SKU"]), int(row["预计数量"]), str(row["是否加急"])))
                    success_count += 1
                
                conn.commit()
                conn.close()
                
                if success_count > 0:
                    st.success(f"✅ 导入成功！共成功导入 {success_count} 条新采购单。")
                    log_action("采购导入", f"成功批量导入了 {success_count} 条采购主数据。")
                if skip_count > 0:
                    st.warning(f"⚠️ 提示：有 {skip_count} 条记录因 PO 号在系统内已存在，被自动跳过。")
                    
        except Exception as e:
            st.error(f"❌ 导入失败，请检查 Excel 格式是否与模板一致！错误原因: {e}")

# --- 岗位三：收货台（无线扫码枪对接窗口） ---
elif role == "📦 收货台（扫码作业）":
    st.title("📦 快递收货无线扫码无线工作台")
    st.write("本界面处于**光标自动聚焦**状态。请直接用扫码枪对准快递面单的【物流单号】进行扫描。")

    # 扫码输入框
    barcode_input = st.text_input("👉 扫码输入框（请确保光标在此处闪烁）", key="pda_scan_field", placeholder="等待扫码枪读取面单...")

    if barcode_input:
        barcode = barcode_input.strip()
        conn = get_db_connection()
        cursor = conn.cursor()
        
        # 1. 在数据库中动态匹配物流单号
        cursor.execute("SELECT po_id, is_urgent, status FROM po_orders WHERE tracking_no = ?", (barcode,))
        match = cursor.fetchone()
        
        if match:
            po_id, is_urgent, current_status = match
            if current_status == "待收货":
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                # 2. 动态更新状态和到仓时间
                cursor.execute("UPDATE po_orders SET status = '已到货', arrival_time = ? WHERE tracking_no = ?", (now_str, barcode))
                conn.commit()
                
                if is_urgent == "是":
                    st.error(f"🚨 【加急单警告】扫码成功！PO单 [{po_id}] 是加急单！系统已全链路高亮，请立刻移交质检组优先处理！")
                    log_action("收货作业", f"🚨 签收加急单！物流单号: {barcode}，PO号: {po_id}")
                else:
                    st.success(f"✅ 【普通单签收】扫码成功！PO单 [{po_id}] 状态已变更为 [已到货]。")
                    log_action("收货作业", f"签收普通单。物流单号: {barcode}，PO号: {po_id}")
            else:
                st.warning(f"⚠️ 提示：物流单号 {barcode} 在系统内已处于【{current_status}】状态，请勿重复扫描！")
        else:
            st.error(f"❌ 警报：系统内未找到物流单号为 [{barcode}] 的采购单！请检查是否属于外来错发件。")
            log_action("异常警告", f"❌ 扫码失败！发现未知物流单号: {barcode}")
            
        conn.close()
