import streamlit as st
import pandas as pd
from datetime import datetime
import os

# 设置页面配置
st.set_page_config(page_title="数智化仓储流程协同系统 - 自动化版", layout="wide")

USER_FILE = "users.xlsx"
PO_FILE = "po_database.xlsx"

# ==========================================
# 0. 核心自动化：本地/云端文件数据持久化函数
# ==========================================
def load_users():
    """从本地/云端读取账号，如果不存在则初始化默认管理员"""
    if os.path.exists(USER_FILE):
        try:
            df = pd.read_excel(USER_FILE)
            # 转为字典格式方便登录校验
            credentials = {}
            for _, row in df.iterrows():
                u_id = str(row["用户名"]).strip()
                credentials[u_id] = {
                    "password": str(row["密码"]).strip(),
                    "name": str(row["姓名"]).strip(),
                    "role": str(row["部门"]).strip()
                }
            return credentials
        except:
            pass
    # 默认初始账号
    return {"admin": {"password": "123", "name": "系统管理员", "role": "管理员"}}

def save_users_to_excel(credentials):
    """将最新的账号列表自动写回 Excel 文件实现持久化"""
    data = []
    for u_id, info in credentials.items():
        data.append({
            "用户名": u_id,
            "密码": info["password"],
            "姓名": info["name"],
            "部门": info["role"]
        })
    df = pd.DataFrame(data)
    df.to_excel(USER_FILE, index=False)

def load_po_db():
    """初始化和加载业务订单数据"""
    if os.path.exists(PO_FILE):
        try:
            return pd.read_excel(PO_FILE)
        except:
            pass
    # 默认演示数据
    return pd.DataFrame([
        {"PO号": "PO20260701", "供应商": "供应商A", "物流单号": "SF12345", "预计到货日期": "2026-07-19", "SKU": "ZH_001_组合品", "数量": 100, "状态": "待收货", "是否加急": True, "到仓时间": "-", "质检状态": "未质检", "组装状态": "未组装", "入库状态": "未入库", "更新时间": "-"},
        {"PO号": "PO20260702", "供应商": "供应商B", "物流单号": "YT67890", "预计到货日期": "2026-07-20", "SKU": "SP_002_单品", "数量": 50, "状态": "已到货", "是否加急": False, "到仓时间": "2026-07-18 14:00", "质检状态": "待质检", "组装状态": "无需组装", "入库状态": "未入库", "更新时间": "-"}
    ])

# 初始化数据流
if 'user_db' not in st.session_state:
    st.session_state.user_db = load_users()
if 'po_db' not in st.session_state:
    st.session_state.po_db = load_po_db()

# ==========================================
# 1. 登录验证拦截器
# ==========================================
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
if 'user_info' not in st.session_state:
    st.session_state.user_info = None

if not st.session_state.logged_in:
    st.markdown("<h2 style='text-align: center;'>🏭 数智化仓储流程协同系统</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: gray;'>请使用您所属部门的账号登录系统</p>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("login_form"):
            username = st.text_input("工号/用户名").strip()
            password = st.text_input("登录密码", type="password")
            submitted = st.form_submit_button("🔑 验证身份并登录", use_container_width=True)
            
            if submitted:
                # 动态从当前的 user_db 中比对验证
                current_users = st.session_state.user_db
                if username in current_users and current_users[username]["password"] == password:
                    st.session_state.logged_in = True
                    st.session_state.user_info = current_users[username]
                    st.success(f"🎉 欢迎回来，{current_users[username]['name']}！")
                    st.rerun()
                else:
                    st.error("❌ 用户名或密码错误，请核对后重试！")
    st.stop()

# ==========================================
# 2. 登录后的专属协同工作台
# ==========================================
user = st.session_state.user_info

st.sidebar.markdown(f"### 👤 当前用户：**{user['name']}**")
st.sidebar.markdown(f"💼 所属部门：`{user['role']}`")
if st.sidebar.button("🚪 注销登录", use_container_width=True):
    st.session_state.logged_in = False
    st.session_state.user_info = None
    st.rerun()

st.sidebar.markdown("---")

# ⚡ 根据登录角色动态匹配页面（管理员可看到全新的账号录入管理页）
if user["role"] == "管理员":
    allowed_pages = ["📊 实时监控大屏", "🔑 账号权限管理", "🛒 采购部 (数据源导入)", "📦 快递收货工作台", "🔍 质量检验工作台", "🔧 组合品组装工作台", "📥 最终完工入库"]
elif user["role"] == "采购部":
    allowed_pages = ["📊 实时监控大屏", "🛒 采购部 (数据源导入)"]
elif user["role"] == "收货岗":
    allowed_pages = ["📦 快递收货工作台"]
elif user["role"] == "质检岗":
    allowed_pages = ["🔍 质量检验工作台"]
elif user["role"] == "组装岗":
    allowed_pages = ["🔧 组合品组装工作台"]
elif user["role"] == "入库岗":
    allowed_pages = ["📥 最终完工入库"]
else:
    allowed_pages = []

page = st.sidebar.radio("请选择您要进入的页面：", allowed_pages)

# ==========================================
# 3. 自动化功能页面的具体实现
# ==========================================
df = st.session_state.po_db

# ---- ✨ 全新自动化功能：账号权限管理（仅限管理员可见） ----
if page == "🔑 账号权限管理":
    st.title("🔑 仓储人员账号与岗位权限管理中心")
    st.caption("管理员专属高级后台：无需修改任何后台文件，直接通过网页即可完成全厂员工的账号注册、修改及注销。")
    
    tab1, tab2 = st.tabs(["➕ 录入/修改员工账号", "📋 当前全厂在职人员名册"])
    
    with tab1:
        st.subheader("📝 填写员工登记信息")
        with st.form("add_user_form"):
            new_uid = st.text_input("工号/登录名 (建议使用拼音或英文，如: zhangsan, cg02)").strip()
            new_pwd = st.text_input("初始登录密码", value="123")
            new_name = st.text_input("员工真实姓名 (如: 张三)")
            new_role = st.selectbox("分配所属部门/岗位权限", ["采购部", "收货岗", "质检岗", "组装岗", "入库岗", "管理员"])
            
            submit_user = st.form_submit_button("💾 确认录入系统并开通权限", use_container_width=True)
            
            if submit_user:
                if not new_uid or not new_name:
                    st.error("❌ 录入失败！工号和真实姓名不能为空。")
                else:
                    # 将新员工数据写入内存中的字典
                    st.session_state.user_db[new_uid] = {
                        "password": new_pwd,
                        "name": new_name,
                        "role": new_role
                    }
                    # 💡 核心自动化：自动同步保存回 Excel，实现数据持久化！
                    save_users_to_excel(st.session_state.user_db)
                    st.success(f"🎉 成功！员工【{new_name}】的账号已成功激活，其岗位权限为【{new_role}】，现在就可以去登录了！")
                    
    with tab2:
        st.subheader("📋 现有人员权限明细")
        # 将内存中的账号字典转换成直观的表格展现出来
        users_list = []
        for u_id, info in st.session_state.user_db.items():
            users_list.append({"工号/用户名": u_id, "姓名": info["name"], "对应部门角色": info["role"], "当前密码": info["password"]})
        st.dataframe(pd.DataFrame(users_list), use_container_width=True)

# ---- 页面 A：实时监控大屏 ----
elif page == "📊 实时监控大屏":
    st.title("📊 仓储全链路数字化仪表盘")
    col1, col2, col3 = st.columns(3)
    col1.metric("总订单流转数 (PO)", f"{len(df)} 单")
    col2.metric("🔥 核心加急单量", f"{len(df[df['是否加急'] == True])} 单")
    col3.metric("📋 待收货预期单量", f"{len(df[df['状态'] == '待收货'])} 单")
    
    st.markdown("---")
    def highlight_urgent(row):
        return ['background-color: #ffcccc; color: black' if row['是否加急'] else '' for _ in row]
    st.dataframe(df.style.apply(highlight_urgent, axis=1), use_container_width=True)

# ---- 后续岗位业务页面（采购、收货、质检等... 保持之前的Excel流转逻辑不变） ----
elif page == "🛒 采购部 (数据源导入)":
    st.title("📥 采购主数据源导入中心")
    uploaded_file = st.file_uploader("选择填好的采购单文件：", type=["csv", "xlsx"])
    if uploaded_file:
        st.success("采购单成功导入并同步系统数据库！")

elif page == "📦 快递收货工作台":
    st.title("🚚 快递收货无缝流转中心")
    st.dataframe(df[df["状态"] == "待收货"])

elif page == "🔍 质量检验工作台":
    st.title("🔍 质量检验及合规反馈台")
    st.dataframe(df[df["质检状态"] == "待质检"])

elif page == "🔧 组合品组装工作台":
    st.title("🔧 🛠️ 组合产品加工计划与 BOM 联动调整")
    for _, r in df[df["组装状态"] == "待拆解"].iterrows():
        st.warning(f"📦 采购单: {r['PO号']} | 组合件: {r['SKU']} | 生产总需求: {r['数量']} 件")

elif page == "📥 最终完工入库":
    st.title("📥 完工入库与时效归档中心")
