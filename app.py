import streamlit as st
import pandas as pd
from datetime import date

from excel_export import generate_payslip_workbook, sample_record
from pdf_export import generate_payslip_pdf

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
    remarks = st.text_input("備考", placeholder="任意")
with col2:
    today_str = date.today().strftime("%Y-%m-%d")
    created_date = st.text_input("作成日", value=today_str, placeholder="YYYY-MM-DD")
    payment_date = st.text_input("支給日", value=today_str, placeholder="YYYY-MM-DD")

# --- 3. 支給項目 ---
st.subheader("支給項目")
base_salary = st.number_input("基本給", min_value=0, value=0, step=1000)
overtime_pay = st.number_input("時間外手当", min_value=0, value=0, step=1000)
family_allowance = st.number_input("家族手当", min_value=0, value=0, step=1000)
commute_allowance = st.number_input("通勤手当", min_value=0, value=0, step=1000)
qualification_allowance = st.number_input("資格手当", min_value=0, value=0, step=1000)
communication_allowance = st.number_input("通信手当", min_value=0, value=0, step=1000)

salary_total_a = (
    base_salary
    + overtime_pay
    + family_allowance
    + commute_allowance
    + qualification_allowance
    + communication_allowance
)
# --- 4. 社会保険 ---
st.subheader("社会保険")
health_insurance = st.number_input("健康保険", min_value=0, value=0, step=1000)
pension = st.number_input("厚生年金", min_value=0, value=0, step=1000)
employment_insurance_rate = st.number_input(
    "雇用保険料率",
    min_value=0.0,
    value=0.005,
    step=0.001,
    format="%.3f",
)
st.caption("例：0.005 = 0.5%")
employment_insurance = round(salary_total_a * employment_insurance_rate)
st.text_input("雇用保険（自動計算）", value=f"{employment_insurance:,}", disabled=True)
childcare_support = st.number_input("子ども・子育て支援金", min_value=0, value=0, step=100)
other_social_insurance = st.number_input("その他社会保険", min_value=0, value=0, step=1000)

# --- 5. 税金・控除 ---
st.subheader("税金・控除")
income_tax = st.number_input("源泉所得税", min_value=0, value=0, step=100)
resident_tax = st.number_input("市町村民税", min_value=0, value=0, step=100)
other_deduction = st.number_input("その他控除", min_value=0, value=0, step=1000)

# 計算
taxable_salary_c = salary_total_a
social_insurance_total_d = (
    health_insurance
    + pension
    + employment_insurance
    + childcare_support
    + other_social_insurance
)
after_social_deduction_e = salary_total_a - social_insurance_total_d
tax_total_f = income_tax + resident_tax + other_deduction
net_payment_g = after_social_deduction_e - tax_total_f

st.divider()

# --- 6. 自動計算結果 ---
st.subheader("自動計算結果")
m1, m2 = st.columns(2)
m1.metric("A 給与総額", f"{salary_total_a:,} 円")
m2.metric("C 課税対象給与", f"{taxable_salary_c:,} 円")
m3, m4 = st.columns(2)
m3.metric("D 社会保険料計", f"{social_insurance_total_d:,} 円")
m4.metric("E 差引控除後給与額", f"{after_social_deduction_e:,} 円")
m5, m6 = st.columns(2)
m5.metric("F 控除計", f"{tax_total_f:,} 円")
m6.metric("G 差引支給額", f"{net_payment_g:,} 円")

if st.button("この内容を記録する", type="primary"):
    if not employee_name.strip():
        st.error("氏名を入力してください。")
    else:
        st.session_state.records.append(
            {
                "作成日": created_date.strip(),
                "支給日": payment_date.strip(),
                "月分": target_month,
                "氏名": employee_name.strip(),
                "備考": remarks.strip(),
                "基本給": base_salary,
                "時間外手当": overtime_pay,
                "家族手当": family_allowance,
                "通勤手当": commute_allowance,
                "資格手当": qualification_allowance,
                "通信手当": communication_allowance,
                "A_給与総額": salary_total_a,
                "C_課税対象給与": taxable_salary_c,
                "健康保険": health_insurance,
                "厚生年金": pension,
                "雇用保険": employment_insurance,
                "子ども・子育て支援金": childcare_support,
                "その他社会保険": other_social_insurance,
                "D_社会保険料計": social_insurance_total_d,
                "E_差引控除後給与額": after_social_deduction_e,
                "源泉所得税": income_tax,
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
    st.dataframe(df, width="stretch")

    # --- 9. ダウンロード ---
    st.subheader("ダウンロード")
    dl_col1, dl_col2, dl_col3 = st.columns(3)

    with dl_col1:
        csv_data = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="CSVでダウンロード（管理・集計用）",
            data=csv_data,
            file_name="salary_records.csv",
            mime="text/csv",
        )

    with dl_col2:
        excel_buffer = generate_payslip_workbook(st.session_state.records)
        st.download_button(
            label="給与明細印刷用Excel（A4・1人1枚）",
            data=excel_buffer.getvalue(),
            file_name="salary_payslips.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with dl_col3:
        pdf_buffer = generate_payslip_pdf(st.session_state.records)
        st.download_button(
            label="給与明細PDFをダウンロード",
            data=pdf_buffer.getvalue(),
            file_name="salary_payslips.pdf",
            mime="application/pdf",
        )

    st.caption("Excel・PDFはA4縦1ページで印刷できます。")

    btn_col1, btn_col2 = st.columns(2)
    with btn_col1:
        if st.button("サンプル1件を記録に追加"):
            st.session_state.records.append(sample_record())
            st.rerun()
    with btn_col2:
        if st.button("記録をクリア"):
            st.session_state.records = []
            st.rerun()
else:
    st.write("まだ記録はありません。")
    if st.button("サンプル1件を記録に追加（印刷プレビュー確認用）"):
        st.session_state.records.append(sample_record())
        st.rerun()

