import pandas as pd
import streamlit as st
import io
import json
import hashlib

# 安全設定
try:
    AUTHORIZED_PASSWORDS = {
        st.secrets["passwords"]["analyst"]: "物料分析師"
    }
except:
    # 如果沒有 secrets 檔案，使用預設密碼
    AUTHORIZED_PASSWORDS = {
        "kanfon2025": "物料分析師"
    }

def check_password():
    """密碼驗證函數"""
    def password_entered():
        entered_password = st.session_state["password"]
        if entered_password in AUTHORIZED_PASSWORDS:
            st.session_state["password_correct"] = True
            st.session_state["user_role"] = AUTHORIZED_PASSWORDS[entered_password]
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # 首次訪問
        st.markdown("#  智慧物料分類工具")
        st.markdown("### 請輸入授權密碼以使用系統")
        st.text_input(
            "授權密碼", 
            type="password", 
            on_change=password_entered, 
            key="password",
            placeholder="請輸入您的專用密碼"
        )
        st.info(" 如需取得使用權限，請聯繫系統管理員")
        return False
    elif not st.session_state["password_correct"]:
        # 密碼錯誤
        st.markdown("#  智慧物料分類工具")
        st.markdown("### 請輸入授權密碼以使用系統")
        st.text_input(
            "授權密碼", 
            type="password", 
            on_change=password_entered, 
            key="password",
            placeholder="請輸入您的專用密碼"
        )
        st.error(" 密碼錯誤，請重新輸入")
        return False
    else:
        # 驗證成功
        st.sidebar.success(f" 歡迎，{st.session_state['user_role']}")
        if st.sidebar.button(" 登出"):
            del st.session_state["password_correct"]
            del st.session_state["user_role"]
            st.rerun()
        return True

# 在主程式開始前檢查密碼
if not check_password():
    st.stop()


@st.cache_data
def load_default_rules():
    return DEFAULT_RULES.copy()

@st.cache_data
def process_excel_file(file_data, sheet_name):
    return pd.read_excel(io.BytesIO(file_data), sheet_name=sheet_name)

# --- 預設分類規則 ---
DEFAULT_RULES = {
    "進口": {
        "condition_type": "currency",
        "rule": "not_ntd",
        "description": "幣別不是NTD的項目"
    },
    "板金": {
        "condition_type": "product_code",
        "rule": "startswith_4KB_and_contains_P",
        "description": "產品編號以4KB開頭且包含P"
    },
    "加工件": {
        "condition_type": "product_code", 
        "rule": "startswith_4KB_contains_MHSLK_or_startswith_kb",
        "description": "產品編號以4KB開頭包含M/H/S/L/K，或以KB開頭"
    },
    "電料": {
        "condition_type": "product_code",
        "rule": "startswith_4KZ", 
        "description": "產品編號以4KZ開頭"
    },
    "市購件": {
        "condition_type": "product_code",
        "rule": "startswith_4SS",
        "description": "產品編號以4SS開頭"
    }
}

# --- 動態規則建立器 ---
def create_custom_rules():
    """讓使用者自訂五大分類的編碼規則"""
    st.subheader(" 五大分類規則設定")
    
    # 固定的五大分類
    FIXED_CATEGORIES = ["進口", "板金", "加工件", "電料", "市購件"]
    
    if 'custom_rules' not in st.session_state:
        st.session_state.custom_rules = DEFAULT_RULES.copy()
    
    # 選擇使用預設規則或自訂規則
    rule_mode = st.radio(
        "選擇分類規則模式：",
        ["使用預設規則", "自訂編碼規則"],
        horizontal=True
    )
    
    if rule_mode == "使用預設規則":
        st.info(" 使用系統預設的分類規則")
        for category, rule_info in DEFAULT_RULES.items():
            st.write(f"**{category}**: {rule_info['description']}")
        return DEFAULT_RULES
    
    else:  # 自訂編碼規則
        st.warning(" 自訂模式：請為每個分類設定編碼規則")
        
        # 為每個固定分類設定規則
        with st.expander(" 編輯五大分類規則", expanded=True):
            
            for category in FIXED_CATEGORIES:
                st.markdown(f"###  {category}")
                
                # 條件類型選擇
                condition_type = st.selectbox(
                    f"選擇 {category} 的判斷條件：",
                    ["product_code", "currency"],
                    key=f"condition_type_{category}",
                    format_func=lambda x: " 產品編號條件" if x == "product_code" else "💱 幣別條件"
                )
                
                if condition_type == "product_code":
                    # 產品編號規則設定
                    rule_type = st.selectbox(
                        f"{category} 的編碼規則：",
                        ["開頭包含", "結尾包含", "包含字串", "不包含", "複合條件"],
                        key=f"rule_type_{category}"
                    )
                    
                    if rule_type == "開頭包含":
                        prefix = st.text_input(
                            f"{category} - 產品編號開頭：", 
                            key=f"prefix_{category}",
                            placeholder="例如：4KB, 4KZ, 4SS"
                        )
                        if prefix:
                            st.session_state.custom_rules[category] = {
                                "condition_type": "product_code",
                                "rule": f"startswith_{prefix}",
                                "description": f"產品編號以 '{prefix}' 開頭"
                            }
                    
                    elif rule_type == "結尾包含":
                        suffix = st.text_input(
                            f"{category} - 產品編號結尾：", 
                            key=f"suffix_{category}",
                            placeholder="例如：-P, _M, -IMP"
                        )
                        if suffix:
                            st.session_state.custom_rules[category] = {
                                "condition_type": "product_code",
                                "rule": f"endswith_{suffix}",
                                "description": f"產品編號以 '{suffix}' 結尾"
                            }
                    
                    elif rule_type == "包含字串":
                        contains = st.text_input(
                            f"{category} - 產品編號包含：", 
                            key=f"contains_{category}",
                            placeholder="例如：P, MOTOR, PCB"
                        )
                        if contains:
                            st.session_state.custom_rules[category] = {
                                "condition_type": "product_code",
                                "rule": f"contains_{contains}",
                                "description": f"產品編號包含 '{contains}'"
                            }
                    
                    elif rule_type == "不包含":
                        not_contains = st.text_input(
                            f"{category} - 產品編號不包含：", 
                            key=f"not_contains_{category}",
                            placeholder="例如：TEMP, TEST"
                        )
                        if not_contains:
                            st.session_state.custom_rules[category] = {
                                "condition_type": "product_code",
                                "rule": f"not_contains_{not_contains}",
                                "description": f"產品編號不包含 '{not_contains}'"
                            }
                    
                    elif rule_type == "複合條件":
                        st.markdown("**複合條件設定：**")
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            prefix_condition = st.text_input(f"{category} - 開頭條件：", key=f"compound_prefix_{category}")
                            contains_condition = st.text_input(f"{category} - 包含條件：", key=f"compound_contains_{category}")
                        
                        with col2:
                            logic_type = st.radio(
                                f"{category} - 邏輯關係：", 
                                ["AND (同時符合)", "OR (任一符合)"], 
                                key=f"logic_{category}"
                            )
                        
                        if prefix_condition and contains_condition:
                            logic = "AND" if "AND" in logic_type else "OR"
                            st.session_state.custom_rules[category] = {
                                "condition_type": "product_code",
                                "rule": f"compound_{logic}_{prefix_condition}_{contains_condition}",
                                "description": f"產品編號開頭 '{prefix_condition}' {logic} 包含 '{contains_condition}'"
                            }
                
                elif condition_type == "currency":
                    # 幣別規則設定
                    currency_rule = st.selectbox(
                        f"{category} 的幣別規則：",
                        ["等於", "不等於", "包含於清單", "不在清單中"],
                        key=f"currency_rule_{category}"
                    )
                    
                    if currency_rule in ["等於", "不等於"]:
                        currency_value = st.text_input(
                            f"{category} - 幣別：", 
                            key=f"currency_value_{category}",
                            placeholder="例如：USD, EUR, JPY, NTD"
                        )
                        if currency_value:
                            rule_prefix = "equals" if currency_rule == "等於" else "not_equals"
                            st.session_state.custom_rules[category] = {
                                "condition_type": "currency",
                                "rule": f"{rule_prefix}_{currency_value.upper()}",
                                "description": f"幣別 {currency_rule} '{currency_value.upper()}'"
                            }
                    
                    else:  # 清單模式
                        currency_list = st.text_input(
                            f"{category} - 幣別清單 (用逗號分隔)：", 
                            key=f"currency_list_{category}",
                            placeholder="例如：USD,EUR,JPY 或 NTD,TWD"
                        )
                        if currency_list:
                            currencies = [c.strip().upper() for c in currency_list.split(',')]
                            rule_prefix = "in_list" if "包含於" in currency_rule else "not_in_list"
                            st.session_state.custom_rules[category] = {
                                "condition_type": "currency",
                                "rule": f"{rule_prefix}_" + ",".join(currencies),
                                "description": f"幣別 {currency_rule}：{', '.join(currencies)}"
                            }
                
                # 顯示目前設定
                if category in st.session_state.custom_rules:
                    current_rule = st.session_state.custom_rules[category]
                    st.success(f"目前規則：{current_rule['description']}")
                else:
                    st.warning("尚未設定規則")
                
                st.divider()
    
        # 新增：規則管理功能
        st.subheader(" 規則管理")
        col1, col2 = st.columns(2)
        
        with col1:
            # 匯出規則
            if st.button("匯出規則設定"):
                rules_json = json.dumps(st.session_state.custom_rules, ensure_ascii=False, indent=2)
                st.download_button(
                    label="下載規則檔案",
                    data=rules_json,
                    file_name="classification_rules.json",
                    mime="application/json"
                )
        
        with col2:
            # 匯入規則
            uploaded_rules = st.file_uploader("匯入規則設定", type=['json'])
            if uploaded_rules is not None:
                try:
                    rules_data = json.load(uploaded_rules)
                    st.session_state.custom_rules = rules_data
                    st.success("規則匯入成功！")
                    st.rerun()
                except Exception as e:
                    st.error(f"規則匯入失敗：{e}")
    return st.session_state.custom_rules

def validate_rules(rules):
    """驗證規則設定的完整性"""
    st.subheader("規則驗證")
    
    FIXED_CATEGORIES = ["進口", "板金", "加工件", "電料", "市購件"]
    missing_rules = []
    incomplete_rules = []
    
    for category in FIXED_CATEGORIES:
        if category not in rules:
            missing_rules.append(category)
        elif not rules[category].get("rule") or rules[category]["rule"] == "custom":
            incomplete_rules.append(category)
    
    if missing_rules:
        st.error(f"以下分類尚未設定規則：{', '.join(missing_rules)}")
        return False
    elif incomplete_rules:
        st.warning(f"以下分類規則不完整：{', '.join(incomplete_rules)}")
        return False
    else:
        st.success("所有分類規則已設定完成")
        # 顯示規則摘要
        with st.expander("規則摘要", expanded=False):
            for category, rule_info in rules.items():
                st.write(f"**{category}**: {rule_info['description']}")
        return True

def test_rules(rules):
    """讓使用者測試分類規則 - 使用與實際分類相同的兩階段邏輯"""
    st.subheader("規則測試")
    
    with st.expander("測試兩階段分類邏輯", expanded=False):
        col1, col2 = st.columns(2)
        
        with col1:
            test_prod_code = st.text_input("測試產品編號：", placeholder="例如：4KB2AAP")
            test_currency = st.text_input("測試幣別：", placeholder="例如：NTD").upper()
        
        with col2:
            if st.button("🔍 測試分類") and test_prod_code:
                st.write(f"**測試產品**：{test_prod_code} / {test_currency}")
                st.write("---")
                
                # 使用與實際分類相同的兩階段邏輯
                result_category = "其他"
                
                # 第一階段：檢查進口
                st.write("**第一階段：檢查進口**")
                if "進口" in rules:
                    import_result = check_rule(test_prod_code, test_currency, rules["進口"])
                    if import_result:
                        st.success("符合進口條件 → **分類：進口**")
                        result_category = "進口"
                    else:
                        st.info("❌ 不符合進口條件 → 繼續檢查國產品分類")
                        
                        # 第二階段：檢查國產品分類
                        st.write("**第二階段：檢查國產品分類**")
                        domestic_categories = ["板金", "加工件", "電料", "市購件"]
                        
                        for category in domestic_categories:
                            if category in rules:
                                result = check_rule(test_prod_code, test_currency, rules[category])
                                status = "✅ 符合" if result else "❌ 不符合"
                                st.write(f"- **{category}**: {status}")
                                st.write(f"  └─ 規則：{rules[category]['description']}")
                                
                                # 顯示板金的詳細檢查
                                if category == "板金" and rules[category]["rule"] == "startswith_4KB_and_contains_P":
                                    condition1 = test_prod_code.startswith("4KB")
                                    condition2 = "P" in test_prod_code
                                    st.write(f"    - 以4KB開頭: {condition1}")
                                    st.write(f"    - 包含字母P: {condition2}")
                                    st.write(f"    - 最終結果: {condition1 and condition2}")
                                
                                # 顯示加工件的詳細檢查
                                if category == "加工件" and rules[category]["rule"] == "startswith_4KB_contains_MHSLK_or_startswith_kb":
                                    condition1 = test_prod_code.startswith("4KB") and any(char in test_prod_code for char in "MHSLK")
                                    condition2 = test_prod_code.startswith("KB") and not test_prod_code.startswith("4KB")
                                    st.write(f"    - 4KB開頭且包含M/H/S/L/K: {condition1}")
                                    st.write(f"    - KB開頭(非4KB): {condition2}")
                                    st.write(f"    - 最終結果: {condition1 or condition2}")
                                
                                if result and result_category == "其他":
                                    result_category = category
                                    st.success(f"**符合條件，分類為：{category}**")
                                    break
                
                st.write("---")
                if result_category != "其他":
                    st.success(f"**最終分類結果：{result_category}**")
                else:
                    st.warning("**最終分類結果：其他**")
# --- 更新的規則檢查函式 ---
def check_rule(prod_code, currency, rule_info):
    """檢查單一規則是否符合 - 支援自訂編碼規則"""
    condition_type = rule_info["condition_type"]
    rule = rule_info["rule"]
    
    if condition_type == "currency":
        if rule == "not_ntd":
            return currency != "NTD" and currency != "NAN" and currency != ""
        elif rule.startswith("equals_"):
            target_currency = rule.replace("equals_", "")
            return currency == target_currency.upper()
        elif rule.startswith("not_equals_"):
            target_currency = rule.replace("not_equals_", "")
            return currency != target_currency.upper()
        elif rule.startswith("in_list_"):
            currency_list = rule.replace("in_list_", "").split(",")
            return currency in [c.upper() for c in currency_list]
        elif rule.startswith("not_in_list_"):
            currency_list = rule.replace("not_in_list_", "").split(",")
            return currency not in [c.upper() for c in currency_list]
    
    elif condition_type == "product_code":
        # 統一轉換為大寫處理
        prod_code_upper = str(prod_code).upper()
        
        # 板金規則：4KB開頭 + 任一位置含P
        if rule == "startswith_4KB_and_contains_P":
            starts_with_4KB = prod_code_upper.startswith("4KB")
            contains_P = "P" in prod_code_upper
            result = starts_with_4KB and contains_P
            
            if hasattr(st.session_state, 'debug_mode') and st.session_state.debug_mode:
                st.write(f"     板金規則詳細檢查：")
                st.write(f"      - 產品編號：`{prod_code}`")
                st.write(f"      - 以4KB開頭：{starts_with_4KB}")
                st.write(f"      - 包含字母P：{contains_P}")
                st.write(f"      - 最終結果：{result}")
            
            return result
            
        # 加工件規則：4KB開頭加特殊字元，或純KB開頭
        elif rule == "startswith_4KB_contains_MHSLK_or_startswith_kb":
            text = str(prod_code).upper()  # 轉為大寫處理
            
            # 條件1：4KB開頭 + 特殊字元
            if text.startswith("4KB"):
                special_chars = "MHSLK"
                return any(char in text for char in special_chars)
            
            # 條件2：KB開頭（但非4KB）
            if text.startswith("KB") and not text.startswith("4KB"):
                return True
            
            return False
        elif rule == "startswith_4KZ":
           return prod_code.upper().startswith("4KZ")
        elif rule == "startswith_4SS":
           return prod_code.upper().startswith("4SS")
        elif rule.startswith("startswith_"):
            prefix = rule.replace("startswith_", "").upper()
            return prod_code_upper.startswith(prefix)
        elif rule.startswith("endswith_"):
            suffix = rule.replace("endswith_", "").upper()
            return prod_code_upper.endswith(suffix)
        elif rule.startswith("contains_"):
            substring = rule.replace("contains_", "").upper()
            return substring in prod_code_upper
        elif rule.startswith("not_contains_"):
            substring = rule.replace("not_contains_", "").upper()
            return substring not in prod_code_upper
        elif rule.startswith("compound_"):
            # 處理複合條件：compound_AND/OR_prefix_contains
            parts = rule.split("_")
            if len(parts) >= 4:
                logic = parts[1].upper()  # AND 或 OR
                prefix_condition = parts[2].upper()
                contains_condition = parts[3].upper()
                
                prefix_match = prod_code_upper.startswith(prefix_condition)
                contains_match = contains_condition in prod_code_upper
                
                if logic == "AND":
                    return prefix_match and contains_match
                elif logic == "OR":
                    return prefix_match or contains_match
                return False  # 如果 logic 不是 AND 或 OR，返回 False
    
    return False

# --- 修改後的分類函式 ---
def assign_main_category_dynamic(df, prod_col, currency_col, rules):
    """
    動態分類函式：根據使用者自訂的規則進行分類
    支援除錯模式，顯示前5筆資料的詳細分類邏輯
    第一階段：檢查是否為進口
    第二階段：對非進口品按產品編號分類
    """
    
    def classify_row(row, show_debug=False):
        try:
            prod_code = str(row[prod_col]).strip() if pd.notna(row[prod_col]) else ""
            currency = str(row[currency_col]).strip().upper() if pd.notna(row[currency_col]) else ""
            
            if show_debug:
                st.markdown(f"---")
                st.write(f"🔍 分析項目: {prod_code}")
                st.write(f"💱 幣別: {currency}")
            # 第一階段：優先檢查是否為進口
            if "進口" in rules:
                import_check = check_rule(prod_code, currency, rules["進口"])
                if show_debug:
                    match_status = "✅" if import_check else "❌"
                    st.write(f"**🌍 第一階段：檢查進口**")
                    st.write(f"{match_status} 檢查 進口 規則: {rules['進口']['description']}")
                
                if import_check:
                    if show_debug:
                        st.success(" 符合進口條件 → **分類：進口**")
                    return "進口"
                else:
                    if show_debug:
                        st.info(" 不符合進口條件 → 繼續檢查國產品分類")
            
            # 第二階段：對非進口品按產品編號進行細分
            if show_debug:
                st.write(f"**第二階段：檢查國產品分類**")
            
            # 確保板金優先判斷
            domestic_categories = ["板金", "加工件", "電料", "市購件"]
            for category in domestic_categories:
                if category in rules:
                    rule_match = check_rule(prod_code, currency, rules[category])
                    
                    if show_debug:
                        match_status = "✅" if rule_match else "❌"
                        st.write(f"- **{category}**: {match_status}")
                        st.write(f"  └─ 規則：{rules[category]['description']}")
                    
                    if rule_match:
                        if show_debug:
                            st.success(f" **符合條件，分類為：{category}**")
                        return category
            
            if show_debug:
                st.warning(" **未符合任何規則，歸類為「其他」**")
            return "其他"
            
        except Exception as e:
            st.warning(f"分類處理錯誤：{e}")
            return "錯誤"
            
    # 檢查除錯模式
    debug_mode = 'debug_mode' in st.session_state and st.session_state.debug_mode
    
    if debug_mode:
        st.write("### 🔍 前5筆資料分類過程")
        # 只對前5筆資料顯示除錯資訊
        for idx, row in df.head().iterrows():
            st.write(f"#### 第 {idx+1} 筆資料")
            df.at[idx, '分類'] = classify_row(row, show_debug=True)
        # 處理剩餘資料
        for idx, row in df.iloc[5:].iterrows():
            df.at[idx, '分類'] = classify_row(row, show_debug=False)
    else:
        # 一般模式：直接應用分類
        df['分類'] = df.apply(lambda row: classify_row(row, show_debug=False), axis=1)
    
    return df

def perform_abc_analysis(df, qty_col, price_col):
    """
    第二階段：在每個主分類內部，獨立進行 ABC 分析
    """
    try:
        # 確保計算欄位是數字，無法轉換的會變成 NaN (Not a Number)
        df['金額'] = pd.to_numeric(df[qty_col], errors='coerce') * pd.to_numeric(df[price_col], errors='coerce')
        
        # 填充可能產生的空值為 0
        df['金額'] = df['金額'].fillna(0)

        # 對 DataFrame 進行排序，以便後續計算累計值
        df = df.sort_values(by=['分類', '金額'], ascending=[True, False])

        # 使用 groupby() 按「分類」分組，並在組內計算累計百分比
        # transform 會返回一個與原 df 同樣大小的 Series，完美解決問題
        df['累計金額'] = df.groupby('分類')['金額'].cumsum()
        group_totals = df.groupby('分類')['金額'].transform('sum')
        df['累計百分比'] = (df['累計金額'] / group_totals).fillna(0)

        # 根據累計百分比分配 ABC 類別 (你的規則是 70/90)
        def assign_abc_category(row):
            if row['金額'] == 0:
                return 'C'
            elif row['累計百分比'] <= 0.7:
                return 'A'
            elif row['累計百分比'] <= 0.9:
                return 'B'
            else:
                return 'C'
        
        # 直接使用函式分配，避免 categorical 問題
        df['ABC類別'] = df.apply(assign_abc_category, axis=1)

        return df
    
    except Exception as e:
        st.error(f"ABC分析過程中發生錯誤：{e}")
        return df
    
# --- 修改後的主介面 ---
st.set_page_config(page_title="智慧物料分類工具", layout="wide")
st.title('智慧物料 ABC 分類工具')
st.write('上傳 Excel，系統將依據您設定的規則進行「主分類」與「ABC 分類」。')

# 步驟1：規則設定
classification_rules = create_custom_rules()

# 步驟1.5：規則驗證
validate_rules(classification_rules)

# 步驟1.6：規則測試
test_rules(classification_rules)

# 新增：除錯模式設定
st.subheader("除錯模式")
st.session_state.debug_mode = st.checkbox("啟用除錯模式（顯示前5筆資料的詳細分類過程）")
if st.session_state.debug_mode:
    st.info("除錯模式已啟用，將在分類時顯示前5筆資料的詳細分類邏輯")

# 步驟2：檔案上傳
st.subheader("檔案上傳")
uploaded_file = st.file_uploader("請選擇你的 Excel 檔案", type=['xlsx'])

if uploaded_file is not None:
    try:
        # 取得工作表
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        
        st.info("偵測到以下工作表，請選擇包含資料的工作表：")
        
        # 改善預設工作表選擇邏輯
        default_sheet_index = 0
        if "uservo2000" in sheet_names:
            default_sheet_index = sheet_names.index("uservo2000")
        elif "分類物料" in sheet_names:
            default_sheet_index = sheet_names.index("分類物料")
        elif len(sheet_names) > 0:
            default_sheet_index = 0

        selected_sheet = st.selectbox(
            label="選擇工作表", 
            options=sheet_names,
            index=default_sheet_index,
        )
        
        if selected_sheet:
            df_original = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
            st.success(f"成功讀取工作表：`{selected_sheet}`！")
            st.dataframe(df_original.head())

        # 欄位選擇
        st.subheader("欄位對應")
        required_cols = {
            "prod_col": "產品編號",
            "currency_col": "幣別", 
            "qty_col": "需求數",
            "price_col": "單價"
        }
        
        all_columns = df_original.columns.tolist()
        
        col1, col2 = st.columns(2)
        with col1:
            prod_col_selected = st.selectbox(f"選擇 '{required_cols['prod_col']}' 對應的欄位:", all_columns, index=1 if len(all_columns) > 1 else 0)
            currency_col_selected = st.selectbox(f"選擇 '{required_cols['currency_col']}' 對應的欄位:", all_columns, index=5 if len(all_columns) > 5 else 0)
        with col2:
            qty_col_selected = st.selectbox(f"選擇 '{required_cols['qty_col']}' 對應的欄位:", all_columns, index=3 if len(all_columns) > 3 else 0)
            price_col_selected = st.selectbox(f"選擇 '{required_cols['price_col']}' 對應的欄位:", all_columns, index=4 if len(all_columns) > 4 else 0)

        # 執行分類
        if st.button(" 開始執行完整分類", type="primary"):
            with st.spinner('正在進行分類，請稍候...'):
                df_processed = df_original.copy()
                
                # 使用動態規則進行分類
                df_processed = assign_main_category_dynamic(df_processed, prod_col_selected, currency_col_selected, classification_rules)
                
                # ABC 分析
                df_final = perform_abc_analysis(df_processed, qty_col_selected, price_col_selected)
                
                st.success("分類完成！")
                
                # 顯示分類統計
                st.subheader("分類統計")
                # 基本統計
                category_stats = df_final['分類'].value_counts()
                abc_stats = df_final['ABC類別'].value_counts()
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.markdown("**主分類統計**")
                    st.bar_chart(category_stats)
                    st.dataframe(category_stats.rename("數量"))
                
                with col2:
                    st.markdown("**ABC分類統計**")
                    st.bar_chart(abc_stats)
                    st.dataframe(abc_stats.rename("數量"))
                
                with col3:
                    st.markdown("**金額統計**")
                    amount_by_category = df_final.groupby('分類')['金額'].sum().sort_values(ascending=False)
                    st.bar_chart(amount_by_category)
                    st.dataframe(amount_by_category.rename("總金額"))
                
                # 新增：交叉分析表
                st.subheader("交叉分析")
                cross_analysis = pd.crosstab(df_final['分類'], df_final['ABC類別'], margins=True)
                st.dataframe(cross_analysis)
                
                # 顯示結果
                st.subheader("分類結果")
                st.dataframe(df_final)

                # 下載功能
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='分類結果')
                
                st.download_button(
                    label="下載分類後的 Excel 檔案",
                    data=output.getvalue(),
                    file_name="classified_materials_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"處理檔案時發生錯誤：{e}")
        st.error("請確認：1. 上傳的是 Excel 檔案。 2. 檔案格式正確。 3. 選擇的欄位正確無誤。")