import base64
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from io import BytesIO
import json
import os
import openpyxl
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from reportlab.graphics.shapes import Drawing, Rect, Line
import smtplib
import streamlit as st

# 引入 Google Drive API 套件 (若未安裝會提示，但不影響主程式運作)
try:
  from google.oauth2 import service_account
  from googleapiclient.discovery import build
  from googleapiclient.http import MediaIoBaseUpload

  GOOGLE_DRIVE_AVAILABLE = True
except ImportError:
  GOOGLE_DRIVE_AVAILABLE = False

# 設定網頁標題與寬度
st.set_page_config(page_title="專業工程驗收單系統", page_icon="🧾", layout="wide")

# 針對手機版面與表格進行響應式優化 CSS
st.markdown(
    """
    <style>
    @media (max-width: 768px) {
        .stDataFrame, div[data-testid="stDataEditor"] {
            width: 100% !important;
            overflow-x: auto;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🧾 專業工程驗收單系統（PDF 與 Excel 雙向驗收管理）")

# ==========================================
# 0. 側邊欄：格式設定與欄寬自訂
# ==========================================
st.sidebar.header("🎨 驗收單樣式與字型設定")

font_options = {
    "微軟正黑體": "C:/Windows/Fonts/msjh.ttc",
    "標楷體": "C:/Windows/Fonts/kaiu.ttf",
    "新細明體": "C:/Windows/Fonts/mingliu.ttc",
}

selected_font_label = st.sidebar.selectbox(
    "選擇 PDF 字型", list(font_options.keys())
)
font_path = font_options[selected_font_label]

if os.path.exists(font_path):
  pdfmetrics.registerFont(TTFont("CUSTOM_FONT", font_path))
  FONT_NAME = "CUSTOM_FONT"
else:
  FONT_NAME = "Helvetica"
  st.sidebar.warning(
      f"找不到字型檔 {font_path}，將自動改用預設英文字型 Helvetica。"
  )

font_size = st.sidebar.slider("內文自訂字型大小", 8, 14, 10, 1)
line_spacing = st.sidebar.slider("內文行距 (Leading)", 10, 20, 14, 1)
title_font_size = st.sidebar.slider("大標題字型大小", 14, 24, 18, 1)
category_font_size = st.sidebar.slider("工作分類名稱字型大小", 10, 20, 14, 1)
total_font_size = st.sidebar.slider("總計金額字型大小", 8, 20, 12, 1)

primary_color_hex = st.sidebar.color_picker("表格標題背景顏色", "#F2F2F2")
text_color_hex = st.sidebar.color_picker("文字顏色", "#000000")

# ------------------------------------------
# 📐 表格欄寬手動調整設定
# ------------------------------------------
st.sidebar.markdown("---")
st.sidebar.header("📐 表格欄寬調整 (總寬需配合A4)")
col_w0 = st.sidebar.slider("驗收 欄寬", 30, 60, 45)
col_w1 = st.sidebar.slider("項次 欄寬", 30, 80, 45)
col_w2 = st.sidebar.slider("工作項目 欄寬", 80, 200, 110)
col_w3 = st.sidebar.slider("單位 欄寬", 30, 80, 40)
col_w4 = st.sidebar.slider("數量 欄寬", 25, 60, 40)
col_w5 = st.sidebar.slider("單價 欄寬", 40, 100, 70)
col_w6 = st.sidebar.slider("複價 欄寬", 40, 100, 70)
col_w7 = st.sidebar.slider("缺失問題 欄寬", 100, 250, 105)

# ------------------------------------------
# ✒️ 簽核欄位排版設定
# ------------------------------------------
st.sidebar.markdown("---")
st.sidebar.header("✒️ 簽核欄位排版調整")
sig_padding_bottom = st.sidebar.slider("簽核欄位列高 (底部留白)", 10, 60, 18)
sig_width_ratio = st.sidebar.slider("簽核欄位寬度比例 (%)", 50, 100, 100)

# ------------------------------------------
# 📧 電子郵件發送設定
# ------------------------------------------
st.sidebar.markdown("---")
st.sidebar.header("📧 電子郵件寄送設定")
smtp_server = st.sidebar.text_input(
    "SMTP 伺服器", "smtp.gmail.com", key="smtp_server_input"
)
smtp_port = st.sidebar.number_input(
    "SMTP 連接埠", value=587, key="smtp_port_input"
)

default_sender = "your_email@gmail.com"
default_pass = ""

try:
  if "email" in st.secrets:
    default_sender = st.secrets["email"].get("sender_email", default_sender)
    default_pass = st.secrets["email"].get("sender_password", default_pass)
except Exception:
  pass

sender_email = st.sidebar.text_input(
    "寄件者信箱", default_sender, key="sender_email_input"
)
sender_password = st.sidebar.text_input(
    "寄件者密碼 (若已設定 Secrets 可留空)",
    value=default_pass,
    type="password",
    key="sender_password_input",
)
receiver_email = st.sidebar.text_input(
    "收件者信箱", "client@example.com", key="receiver_email_input"
)

# ------------------------------------------
# ☁️ Google 雲端硬碟設定
# ------------------------------------------
st.sidebar.markdown("---")
st.sidebar.header("☁️ 傳送到 Google 雲端硬碟")
gdrive_enabled = st.sidebar.checkbox("啟用 Google 雲端上傳功能")
uploaded_key_file = st.sidebar.file_uploader(
    "Google 服務帳戶金鑰", type=["json"]
)


# ==========================================
# 初始化 Session State
# ==========================================
if "categories" not in st.session_state:
  st.session_state["categories"] = [
      {
          "name": "一、 植栽換植工程",
          "data": pd.DataFrame([
              {
                  "驗收": False,
                  "項次": "1-1",
                  "工作項目": "植栽清運",
                  "單位": "式",
                  "數量": 1,
                  "單價(元)": 10000,
                  "複價(元)": 10000,
                  "缺失問題": "",
              },
              {
                  "驗收": False,
                  "項次": "1-2",
                  "工作項目": "長虹木",
                  "單位": "棵",
                  "數量": 10,
                  "單價(元)": 250,
                  "複價(元)": 2500,
                  "缺失問題": "",
              },
              {
                  "驗收": False,
                  "項次": "1-3",
                  "工作項目": "土壤改良資材",
                  "單位": "式",
                  "數量": 1,
                  "單價(元)": 1000,
                  "複價(元)": 1000,
                  "缺失問題": "",
              },
              {
                  "驗收": False,
                  "項次": "1-4",
                  "工作項目": "植栽工",
                  "單位": "工",
                  "數量": 2,
                  "單價(元)": 2800,
                  "複價(元)": 5600,
                  "缺失問題": "",
              },
              {
                  "驗收": False,
                  "項次": "1-5",
                  "工作項目": "植栽運費",
                  "單位": "趟",
                  "數量": 1,
                  "單價(元)": 1000,
                  "複價(元)": 1000,
                  "缺失問題": "",
              },
          ]),
          "history": [],
      }
  ]


# ==========================================
# 0-1. 匯入 Excel 功能區塊
# ==========================================
st.markdown("### 📊 匯入 Excel 驗收單")
uploaded_excel = st.file_uploader(
    "選擇 Excel 檔案 (.xlsx / .xls)", type=["xlsx", "xls"], key="excel_uploader"
)
if uploaded_excel is not None:
  try:
    excel_file = pd.ExcelFile(uploaded_excel)
    df_imported = pd.read_excel(uploaded_excel, sheet_name=0)

    parsed_categories = []
    current_cat_name = "一、 匯入工作項目"
    current_rows = []

    required_cols = [
        "驗收",
        "項次",
        "工作項目",
        "單位",
        "數量",
        "單價(元)",
        "複價(元)",
        "缺失問題",
    ]

    for idx, row in df_imported.iterrows():
      row_vals = [str(val).strip() if pd.notna(val) else "" for val in row.values]
      val_str = row_vals[0] if len(row_vals) > 0 else ""

      is_category_row = False
      if ("、" in val_str or "工程" in val_str or "項目" in val_str) and (
          len(row_vals) <= 2 or all(v == "" for v in row_vals[2:])
      ):
        is_category_row = True

      if is_category_row:
        if current_rows:
          parsed_categories.append({
              "name": current_cat_name,
              "data": pd.DataFrame(current_rows, columns=required_cols),
              "history": [],
          })
          current_rows = []
        current_cat_name = val_str
      else:
        if any(row_vals):
          raw_check = (
              row_vals[0]
              if len(row_vals) > 0 and str(row_vals[0]).lower() == "true"
              else False
          )
          item_data = {
              "驗收": bool(raw_check),
              "項次": row_vals[1] if len(row_vals) > 1 else "",
              "工作項目": row_vals[2] if len(row_vals) > 2 else "",
              "單位": row_vals[3] if len(row_vals) > 3 else "式",
              "數量": (
                  pd.to_numeric(row_vals[4], errors="coerce")
                  if len(row_vals) > 4 and row_vals[4] != ""
                  else 1
              ),
              "單價(元)": (
                  pd.to_numeric(row_vals[5], errors="coerce")
                  if len(row_vals) > 5 and row_vals[5] != ""
                  else 0
              ),
              "複價(元)": (
                  pd.to_numeric(row_vals[6], errors="coerce")
                  if len(row_vals) > 6 and row_vals[6] != ""
                  else 0
              ),
              "缺失問題": row_vals[7] if len(row_vals) > 7 else "",
          }
          item_data["複價(元)"] = item_data["數量"] * item_data["單價(元)"]
          current_rows.append(item_data)

    if current_rows:
      parsed_categories.append({
          "name": current_cat_name,
          "data": pd.DataFrame(current_rows, columns=required_cols),
          "history": [],
      })

    if len(parsed_categories) == 0:
      df_clean = df_imported.copy()
      for col in required_cols:
        if col not in df_clean.columns:
          if col == "驗收":
            df_clean[col] = False
          else:
            df_clean[col] = ""
      df_clean = df_clean[required_cols]
      parsed_categories = [{
          "name": "一、 匯入工作項目",
          "data": df_clean,
          "history": [],
      }]

    st.session_state["categories"] = parsed_categories
    st.success(
        f"成功匯入 Excel 檔案！共自動識別並載入 {len(parsed_categories)} 個大項工項。"
    )
    st.rerun()
  except Exception as e:
    st.error(f"匯入 Excel 失敗，請確認檔案格式是否正確：{e}")

st.markdown("---")


# ==========================================
# 1. 基本資訊與工程標頭
# ==========================================
st.subheader("1. 工程與客戶基本資訊")
col1, col2 = st.columns(2)

with col1:
  company_title = st.text_input("公司/標題名稱", "OO公司驗收單")
  project_name = st.text_input("工程名稱", "OO工程")
  client_name = st.text_input("會員/業主名稱", "OOO")

with col2:
  location = st.text_input("施工地點", "地址")
  construction_date = st.date_input("施工日期")
  acceptance_date = st.date_input("驗收日期")
  client_tax_id = st.text_input("統一編號", "8碼")


# ==========================================
# 1-1. 簽核欄自訂與勾選設定
# ==========================================
st.subheader("1-1. 簽核欄位勾選與自訂")
sig_col1, sig_col2, sig_col3, sig_col4 = st.columns(4)

with sig_col1:
  enable_sig1 = st.checkbox("啟用簽核 1", value=True)
  sig_label1 = st.text_input("簽核 1 名稱", "主管簽核:")
with sig_col2:
  enable_sig2 = st.checkbox("啟用簽核 2", value=True)
  sig_label2 = st.text_input("簽核 2 名稱", "")
with sig_col3:
  enable_sig3 = st.checkbox("啟用簽核 3", value=True)
  sig_label3 = st.text_input("簽核 3 名稱", "客戶確認:")
with sig_col4:
  enable_sig4 = st.checkbox("啟用簽核 4", value=True)
  sig_label4 = st.text_input("簽核 4 名稱", "製表人/經辦人:")

active_signatures = []
if enable_sig1:
  active_signatures.append(sig_label1)
if enable_sig2:
  active_signatures.append(sig_label2)
if enable_sig3:
  active_signatures.append(sig_label3)
if enable_sig4:
  active_signatures.append(sig_label4)


# ==========================================
# 2. 動態大項工項與明細編輯區
# ==========================================
st.markdown("---")
sec_col, calc_col = st.columns([4, 1])
with sec_col:
  st.subheader("2. 大項工項與明細編輯")

all_units_set = {"式", "棵", "工", "趟", "個", "項", "m", "m²", "m³", "批", "組"}
all_items_set = {
    "植栽清運",
    "長虹木",
    "土壤改良資材",
    "植栽工",
    "植栽運費",
}

for cat_item in st.session_state["categories"]:
  if "單位" in cat_item["data"].columns:
    for u in cat_item["data"]["單位"].dropna():
      if str(u).strip():
        all_units_set.add(str(u).strip())
  if "工作項目" in cat_item["data"].columns:
    for itm in cat_item["data"]["工作項目"].dropna():
      if str(itm).strip():
        all_items_set.add(str(itm).strip())

unit_options = list(all_units_set)
item_options = list(all_items_set)

column_config = {
    "驗收": st.column_config.CheckboxColumn("驗收", default=False),
    "單位": st.column_config.SelectboxColumn(
        "單位", options=unit_options, default="式", required=False
    ),
    "工作項目": st.column_config.SelectboxColumn(
        "工作項目", options=item_options, default="", required=False
    ),
    "複價(元)": st.column_config.NumberColumn("複價(元)", disabled=True),
}

subtotal = 0
category_chinese_nums = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]

indices_to_remove = []

for idx, cat in enumerate(st.session_state["categories"]):
  if "history" not in cat:
    cat["history"] = []

  col_title, col_undo, col_clear, col_remove = st.columns([4, 1.2, 1.2, 1.2])

  with col_title:
    st.markdown(f"### 第 {idx+1} 組大項工項")

  with col_undo:
    if st.button("↩️ 復原", key=f"undo_cat_{idx}"):
      if cat["history"]:
        cat["data"] = cat["history"].pop()
        st.rerun()

  with col_clear:
    if st.button("🧹 清空", key=f"clear_cat_{idx}"):
      cat["history"].append(cat["data"].copy())
      st.session_state["categories"][idx]["data"] = pd.DataFrame(
          columns=[
              "驗收",
              "項次",
              "工作項目",
              "單位",
              "數量",
              "單價(元)",
              "複價(元)",
              "缺失問題",
          ]
      )
      st.rerun()

  with col_remove:
    if idx > 0:
      if st.button("❌ 移除", key=f"remove_cat_{idx}"):
        indices_to_remove.append(idx)

  default_name = cat["name"]
  cat["name"] = st.text_input(
      f"工作分類名稱 {idx+1}", default_name, key=f"cat_name_{idx}"
  )
old_df = cat["data"].copy()

edited_df = st.data_editor(
      cat["data"],
      num_rows="dynamic",
      width="stretch",
      key=f"df_cat_{idx}",
      column_config=column_config,
  )

  # 自動同步編輯後的資料與計算複價，只要內容有變動就直接儲存並重新整理
if not edited_df.equals(old_df):
      edited_df["複價(元)"] = pd.to_numeric(
          edited_df["數量"], errors="coerce"
      ).fillna(0) * pd.to_numeric(
          edited_df["單價(元)"], errors="coerce"
      ).fillna(0)
      cat["history"].append(old_df)
      cat["data"] = edited_df.copy()
      st.rerun()

if not cat["data"].empty and "複價(元)" in cat["data"]:
      subtotal += (
          pd.to_numeric(cat["data"]["複價(元)"], errors="coerce").fillna(0).sum()
      )

if idx == 0:
    st.markdown("")
    if st.button("➕ 新增下一項大項"):
      next_idx = len(st.session_state["categories"])
      prefix = (
          category_chinese_nums[next_idx]
          if next_idx < len(category_chinese_nums)
          else str(next_idx + 1)
      )
      new_cat_name = f"{prefix}、 新增工作項目"
      new_df = pd.DataFrame([
          {
              "驗收": False,
              "項次": f"{next_idx+1}-1",
              "工作項目": "",
              "單位": "式",
              "數量": 1,
              "單價(元)": 0,
              "複價(元)": 0,
              "缺失問題": "",
          }
      ])
      st.session_state["categories"].append({
          "name": new_cat_name,
          "data": new_df,
          "history": [],
      })
      st.rerun()

st.markdown("---")

if indices_to_remove:
  for i in sorted(indices_to_remove, reverse=True):
    st.session_state["categories"].pop(i)
  st.rerun()


# ==========================================
# 3. 金額計算與備註
# ==========================================
tax_rate = 5
tax_amount = subtotal * (tax_rate / 100)
total_amount = subtotal + tax_amount

st.info(f"📊 **目前複價總金額統計：** $ {subtotal:,.0f} 元（尚未含 5% 稅金）")

default_remarks = ""
remarks = st.text_area("備註事項", default_remarks, height=120)


# ==========================================
# 輔助函數：建立完美的 SVG 勾選方框 (解決中文字型無法顯示勾選符號的問題)
# ==========================================
def get_checkbox_drawing(is_checked: bool):
  d = Drawing(12, 12)
  # 繪製外框正方形
  d.add(Rect(0, 0, 12, 12, fillColor=colors.white, strokeColor=colors.black, strokeWidth=1, rx=1, ry=1))
  if is_checked:
    # 繪製打勾的兩條線 (形成 ☑ 樣式)
    d.add(Line(2, 6, 5, 2, strokeColor=colors.black, strokeWidth=1.5))
    d.add(Line(5, 2, 10, 10, strokeColor=colors.black, strokeWidth=1.5))
  return d


# ==========================================
# 4. 使用 ReportLab 產生 PDF 的函數
# ==========================================
def generate_reportlab_pdf():
  buffer = BytesIO()
  doc = SimpleDocTemplate(
      buffer,
      pagesize=A4,
      rightMargin=30,
      leftMargin=30,
      topMargin=30,
      bottomMargin=30,
  )
  elements = []

  from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet

  styles = getSampleStyleSheet()

  txt_color = colors.HexColor(text_color_hex)
  bg_color = colors.HexColor(primary_color_hex)

  title_style = ParagraphStyle(
      "TitleStyle",
      fontName=FONT_NAME,
      fontSize=title_font_size,
      alignment=1,
      spaceAfter=10,
      textColor=txt_color,
  )
  category_style = ParagraphStyle(
      "CategoryStyle",
      fontName=FONT_NAME,
      fontSize=category_font_size,
      leading=category_font_size + 4,
      textColor=txt_color,
  )
  normal_style = ParagraphStyle(
      "NormalStyle",
      fontName=FONT_NAME,
      fontSize=font_size,
      leading=line_spacing,
      textColor=txt_color,
  )
  center_bold_style = ParagraphStyle(
      "CenterBoldStyle",
      fontName=FONT_NAME,
      fontSize=font_size,
      leading=line_spacing,
      alignment=1,
      textColor=txt_color,
  )
  total_style = ParagraphStyle(
      "TotalStyle",
      fontName=FONT_NAME,
      fontSize=total_font_size,
      leading=total_font_size + 4,
      alignment=2,
      textColor=txt_color,
  )

  elements.append(Paragraph(company_title, title_style))

  c_date_str = construction_date.strftime("%Y年%m月%d日")
  a_date_str = acceptance_date.strftime("%Y年%m月%d日")
  
  # 需求 1：將驗收日期移到第二列，在施工日期後方
  meta_text = f"""
    <b>工程名稱：</b> {project_name} &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <b>施工地點：</b> {location}<br/>
    <b>會員/業主：</b> {client_name} (統編:{client_tax_id}) &nbsp;&nbsp;&nbsp;&nbsp; <b>施工日期：</b> {c_date_str} &nbsp;&nbsp; <b>驗收日期：</b> {a_date_str}
    """
  elements.append(Paragraph(meta_text, normal_style))
  elements.append(Spacer(1, 10))

  table_data = [
      [
          Paragraph("<b>驗收<br/>CHECK</b>", center_bold_style),
          Paragraph("<b>項次<br/>ITEM</b>", center_bold_style),
          Paragraph("<b>工作項目<br/>DESCRIPTION</b>", center_bold_style),
          Paragraph("<b>單位<br/>UNIT</b>", center_bold_style),
          Paragraph("<b>數量<br/>Q'TY</b>", center_bold_style),
          Paragraph("<b>單價(元)<br/>UNIT PRICE</b>", center_bold_style),
          Paragraph("<b>複價(元)<br/>AMOUNT</b>", center_bold_style),
          Paragraph("<b>缺失問題<br/>REMARKS</b>", center_bold_style),
      ]
  ]

  span_rows_categories = []

  for cat in st.session_state["categories"]:
    cat_idx = len(table_data)
    table_data.append(
        [Paragraph(f"<b>{cat['name']}</b>", category_style)] + [""] * 7
    )
    span_rows_categories.append(cat_idx)

    for _, row in cat["data"].iterrows():
      is_checked = bool(row.get("驗收", False))
      # 需求 2：使用 SVG 繪圖取代字型字元，確保 PDF 完美呈現勾選框圖示
      checkbox_element = get_checkbox_drawing(is_checked)

      table_data.append([
          checkbox_element,
          Paragraph(str(row.get("項次", "")), center_bold_style),
          Paragraph(str(row.get("工作項目", "")), normal_style),
          Paragraph(str(row.get("單位", "")), center_bold_style),
          Paragraph(str(row.get("數量", 0)), center_bold_style),
          Paragraph(f"{int(row.get('單價(元)', 0)):,}", normal_style),
          Paragraph(f"{int(row.get('複價(元)', 0)):,}", normal_style),
          Paragraph(str(row.get("缺失問題", "")), normal_style),
      ])

  table_data.append([
      Paragraph("<b>小計</b>", normal_style),
      "",
      "",
      "",
      "",
      "",
      Paragraph(f"<b>{subtotal:,.0f}</b>", normal_style),
      "",
  ])
  table_data.append([
      Paragraph("<b>5%稅金</b>", normal_style),
      "",
      "",
      "",
      "",
      "",
      Paragraph(f"<b>{tax_amount:,.0f}</b>", normal_style),
      "",
  ])
  table_data.append([
      Paragraph("<b>總計</b>", normal_style),
      "",
      "",
      "",
      "",
      "",
      Paragraph(f"<b>$ &nbsp; {total_amount:,.0f}</b>", total_style),
      "",
  ])
  table_data.append([
      Paragraph("<b>備註：</b>", normal_style),
      Paragraph(remarks.replace("\n", "<br/>"), normal_style),
      "1.驗收完成日期:",
      "",
      "",
      "",
      "",
      "",
  ])

  t_styles = [
      ("BACKGROUND", (0, 0), (-1, 0), bg_color),
      ("ALIGN", (0, 0), (-1, -1), "CENTER"),
      ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
      ("GRID", (0, 0), (-1, -1), 1, colors.black),
      ("SPAN", (0, -4), (5, -4)),
      ("SPAN", (0, -3), (5, -3)),
      ("SPAN", (0, -2), (5, -2)),
      ("SPAN", (1, -1), (7, -1)),
  ]

  for c_idx in span_rows_categories:
    t_styles.append(("SPAN", (0, c_idx), (-1, c_idx)))

  total_table_widths = [
      col_w0,
      col_w1,
      col_w2,
      col_w3,
      col_w4,
      col_w5,
      col_w6,
      col_w7,
  ]

  t = Table(table_data, colWidths=total_table_widths, repeatRows=1)
  t.setStyle(TableStyle(t_styles))
  elements.append(t)
  elements.append(Spacer(1, 10))

  if active_signatures:
    n_sigs = len(active_signatures)
    total_width = sum(total_table_widths) * (sig_width_ratio / 100.0)
    sig_col_width = total_width / n_sigs
    sig_col_widths = [sig_col_width] * n_sigs

    sig_row = [
        Paragraph(sig_text, normal_style) for sig_text in active_signatures
    ]
    sig_table_data = [sig_row]

    sig_styles = [
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 1, colors.black),
        ("BOTTOMPADDING", (0, 0), (-1, -1), sig_padding_bottom),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
    ]

    sig_t = Table(sig_table_data, colWidths=sig_col_widths)
    sig_t.setStyle(TableStyle(sig_styles))
    elements.append(sig_t)

  doc.build(elements)
  buffer.seek(0)
  return buffer.getvalue()


# ==========================================
# 5. 匯出、寄信與上傳雲端操作區塊
# ==========================================
st.markdown("---")
st.subheader("💾 PDF 匯出、郵件寄送與雲端備份")

if st.button("📄 產生 PDF 檔案", type="primary"):
  pdf_data = generate_reportlab_pdf()
  st.session_state["pdf_data"] = pdf_data
  st.success("PDF 驗收單產生成功！")

if "pdf_data" in st.session_state:
  file_name_str = f"{project_name}_驗收單.pdf"

  if "show_preview" not in st.session_state:
    st.session_state["show_preview"] = False

  col_dl, col_prev = st.columns([1, 1])
  with col_dl:
    st.download_button(
        label="📥 下載正式 PDF 驗收單",
        data=st.session_state["pdf_data"],
        file_name=file_name_str,
        mime="application/pdf",
    )
  with col_prev:
    if st.button("👀 預覽 PDF 驗收單"):
      st.session_state["show_preview"] = not st.session_state["show_preview"]

  if st.session_state["show_preview"]:
    st.markdown("### 🔍 PDF 預覽畫面")
    # 需求 3：針對手機內嵌預覽限制，提供友善提示，同時嘗試利用標準 iframe，若手機瀏覽器不支援則建議直接點選上方下載按鈕
    st.info("💡 提示：若您的手機瀏覽器無法直接預覽 PDF，請直接點選上方「下載正式 PDF 驗收單」進行檢視。")
    base64_pdf = base64.b64encode(st.session_state["pdf_data"]).decode("utf-8")
    pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="700px" type="application/pdf"></iframe>'
    st.markdown(pdf_display, unsafe_allow_html=True)
    st.markdown("---")

  col_email, col_gdrive = st.columns(2)

  with col_email:
    if st.button("📧 發送電子郵件給業主"):
      try:
        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = receiver_email
        msg["Subject"] = f"【{company_title}】{project_name} 驗收單"

        body = (
            f"您好：\n\n附件為「{project_name}」之工程驗收單 PDF 檔案，請查收。\n謝謝！"
        )
        msg.attach(MIMEText(body, "plain"))

        part = MIMEBase("application", "octet-stream")
        part.set_payload(st.session_state["pdf_data"])
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition", f"attachment; filename= {file_name_str}"
        )
        msg.attach(part)

        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receiver_email, msg.as_string())
        server.quit()
        st.success(f"成功將郵件發送至 {receiver_email}！")
      except Exception as e:
        st.error(f"郵件發送失敗，請檢查設定與密碼：{e}")

  with col_gdrive:
    if st.button("☁️ 上傳至 Google 雲端硬碟"):
      if not gdrive_enabled:
        st.warning("請先在側邊欄勾選「啟用 Google 雲端上傳功能」")
      elif not uploaded_key_file:
        st.warning("請在側邊欄上傳您的 Google 服務帳戶 JSON 金鑰檔案")
      elif not GOOGLE_DRIVE_AVAILABLE:
        st.error(
            "未安裝 Google API 套件，請先在終端機執行 `pip install"
            " google-api-python-client google-auth`"
        )
      else:
        try:
          key_data = json.load(uploaded_key_file)
          SCOPES = ["https://www.googleapis.com/auth/drive.file"]
          creds = service_account.Credentials.from_service_account_info(
              key_data, scopes=SCOPES
          )
          service = build("drive", "v3", credentials=creds)

          file_metadata = {"name": file_name_str}
          media = MediaIoBaseUpload(
              BytesIO(st.session_state["pdf_data"]),
              mimetype="application/pdf",
              resumable=True,
          )
          file = (
              service.files()
              .create(
                  body=file_metadata,
                  media_body=media if "media" in locals() else media,
                  fields="id",
              )
              .execute()
          )
          st.success(
              f"成功上傳至 Google 雲端硬碟！檔案 ID: {file.get('id')}"
          )
        except Exception as e:
          st.error(f"上傳至 Google 雲端硬碟失敗：{e}")


# ==========================================
# 6. 下載 Excel 檔
# ==========================================
st.markdown("---")
st.subheader("📊 匯出 Excel 驗收單")

output_excel_buffer = BytesIO()
with pd.ExcelWriter(output_excel_buffer, engine="openpyxl") as writer:
  export_rows = []
  for cat in st.session_state["categories"]:
    export_rows.append({
        "驗收": "",
        "項次": cat["name"],
        "工作項目": "",
        "單位": "",
        "數量": "",
        "單價(元)": "",
        "複價(元)": "",
        "缺失問題": "",
    })
    for _, r in cat["data"].iterrows():
      export_rows.append({
          "驗收": r.get("驗收", False),
          "項次": r.get("項次", ""),
          "工作項目": r.get("工作項目", ""),
          "單位": r.get("單位", ""),
          "數量": r.get("數量", 0),
          "單價(元)": r.get("單價(元)", 0),
          "複價(元)": r.get("複價(元)", 0),
          "缺失問題": r.get("缺失問題", ""),
      })
  df_export = pd.DataFrame(export_rows)
  df_export.to_excel(writer, sheet_name="工程驗收單", index=False)

output_excel_buffer.seek(0)
st.download_button(
    label="📥 下載 Excel 驗收單 (.xlsx)",
    data=output_excel_buffer,
    file_name=f"{project_name}_驗收單.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)