import streamlit as st
import pandas as pd
from datetime import date

st.set_page_config(page_title="給与明細かんたん計算", layout="centered")

# --- 1. タイトル ---
st.title("給与明細かんたん計算")
st.caption("紙の給与明細の数字を入力し、A〜Gを自動計算します。")

if "records" not in st.session_state:
    st.session_state.records = []

# --- 2. 基本情報 ---
st.subheader("基本情報")
col1, col2 = st.columns(2)
with col1:
    target_month = st.text_input("月分", placeholder="例: 4月分")
    employee_name = st.text_input("氏名")
with col2:
    created_date = st.date_input("作成日", value=date.today())

# --- 3. 支給項目 ---
st.subheader("支給項目")
base_salary = st.number_input("基本給", min_value=0, value=0, step=1000)
overtime_pay = st.number_input("時間外手当", min_value=0, value=0, step=1000)
commute_pay = st.number_input("通勤手当", min_value=0, value=0, step=1000)
other_pay_1 = st.number_input("その他手当1", min_value=0, value=0, step=1000)
other_pay_2 = st.number_input("その他手当2", min_value=0, value=0, step=1000)
other_pay_3 = st.number_input("その他手当3", min_value=0, value=0, step=1000)

# --- 4. 非課税・課税対象 ---
st.subheader("非課税・課税対象")
non_tax_commute = st.number_input("非課税通勤", min_value=0, value=0, step=1000)

# --- 5. 社会保険 ---
st.subheader("社会保険")
health_insurance = st.number_input("健康保険", min_value=0, value=0, step=1000)
pension = st.number_input("厚生年金", min_value=0, value=0, step=1000)
employment_insurance = st.number_input("雇用保険", min_value=0, value=0, step=100)
other_social_insurance = st.number_input("その他社会保険", min_value=0, value=0, step=1000)

# --- 6. 税・控除 ---
st.subheader("税・控除")
income_tax = st.number_input("所得税", min_value=0, value=0, step=100)
resident_tax = st.number_input("市町村民税", min_value=0, value=0, step=100)
other_deduction = st.number_input("その他控除", min_value=0, value=0, step=1000)

# 計算
salary_total_a = (
    base_salary + overtime_pay + commute_pay + other_pay_1 + other_pay_2 + other_pay_3
)
non_tax_commute_b = non_tax_commute
taxable_salary_c = salary_total_a - non_tax_commute_b
social_insurance_total_d = (
    health_insurance + pension + employment_insurance + other_social_insurance
)
after_social_deduction_e = salary_total_a - social_insurance_total_d
tax_total_f = income_tax + resident_tax + other_deduction
net_payment_g = after_social_deduction_e - tax_total_f

st.divider()

# --- 7. 自動計算結果 ---
st.subheader("自動計算結果")
m1, m2 = st.columns(2)
m1.metric("A 給与総額", f"{salary_total_a:,} 円")
m2.metric("B 非課税通勤", f"{non_tax_commute_b:,} 円")
m3, m4 = st.columns(2)
m3.metric("C 課税対象給与", f"{taxable_salary_c:,} 円")
m4.metric("D 社会保険料計", f"{social_insurance_total_d:,} 円")
m5, m6 = st.columns(2)
m5.metric("E 差引控除後給与額", f"{after_social_deduction_e:,} 円")
m6.metric("F 控除計", f"{tax_total_f:,} 円")
st.metric("G 差引支給額", f"{net_payment_g:,} 円")

if st.button("この内容を記録する", type="primary"):
    if not employee_name.strip():
        st.error("氏名を入力してください。")
    else:
        st.session_state.records.append(
            {
                "作成日": created_date.strftime("%Y-%m-%d"),
                "月分": target_month,
                "氏名": employee_name.strip(),
                "基本給": base_salary,
                "時間外手当": overtime_pay,
                "通勤手当": commute_pay,
                "その他手当1": other_pay_1,
                "その他手当2": other_pay_2,
                "その他手当3": other_pay_3,
                "A_給与総額": salary_total_a,
                "B_非課税通勤": non_tax_commute_b,
                "C_課税対象給与": taxable_salary_c,
                "健康保険": health_insurance,
                "厚生年金": pension,
                "雇用保険": employment_insurance,
                "その他社会保険": other_social_insurance,
                "D_社会保険料計": social_insurance_total_d,
                "E_差引控除後給与額": after_social_deduction_e,
                "所得税": income_tax,
                "市町村民税": resident_tax,
                "その他控除": other_deduction,
                "F_控除計": tax_total_f,
                "G_差引支給額": net_payment_g,
            }
        )
        st.success("記録しました。")

st.divider()

# --- 8. 記録一覧 ---
st.subheader("記録一覧")

if st.session_state.records:
    df = pd.DataFrame(st.session_state.records)
    st.dataframe(df, use_container_width=True)

    # --- 9. CSVダウンロード ---
    st.subheader("CSVダウンロード")
    csv_data = df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label="CSVでダウンロード",
        data=csv_data,
        file_name="salary_records.csv",
        mime="text/csv",
    )

    if st.button("記録をクリア"):
        st.session_state.records = []
        st.rerun()
else:
    st.write("まだ記録はありません。")

st.divider()

# --- 10. 注意書き ---
st.info(
    "※このアプリは社内確認用の簡易計算ツールです。"
    "正式な給与計算・税額・社会保険料は会社の規定に従って確認してください。"
)
