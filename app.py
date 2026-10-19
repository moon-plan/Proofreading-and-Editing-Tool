import io
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from hanspell import spell_checker

st.set_page_config(
    page_title="맞춤법 및 교정교열 에디터", page_icon="📝", layout="wide"
)

st.title("📝 프로젝트 맞춤형 교정교열 프로그램")
st.markdown(
    "PDF 문서와 프로젝트별 기준 엑셀 파일을 업로드하면, 부산대 맞춤법 검사기 엔진과 커스텀 규칙을 대조하여 검수 리포트를 생성합니다."
)

st.sidebar.header("📂 파일 업로드")
uploaded_pdf = st.sidebar.file_uploader(
    "1. 교정할 PDF 파일 업로드", type=["pdf"]
)
uploaded_excel = st.sidebar.file_uploader(
    "2. 프로젝트별 기준 엑셀 업로드 (.xlsx)", type=["xlsx"]
)


def load_custom_rules(excel_file):
  try:
    df = pd.read_excel(excel_file)
    whitelist = set(
        df[df["적용 예외"] == True]["검출 단어"].dropna().tolist()
        if "적용 예외" in df.columns
        else []
    )
    replacements = (
        dict(
            zip(df["검출 단어"].dropna(), df["올바른 표기"].dropna())
        )
        if "올바른 표기" in df.columns
        else {}
    )
    return whitelist, replacements
  except Exception as e:
    st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")
    return set(), {}


def extract_text_from_pdf(pdf_bytes):
  doc = fitz.open(stream=pdf_bytes, filetype="pdf")
  extracted = []
  for page_num, page in enumerate(doc, start=1):
    text = page.get_text()
    lines = [line.strip() for line in text.split("\n") if line.strip()]
    for line_num, line in enumerate(lines, start=1):
      extracted.append(
          {"page": page_num, "line_no": line_num, "text": line}
      )
  return extracted


if st.button("🚀 교정교열 검수 시작하기", type="primary"):
  if not uploaded_pdf:
    st.warning("교정할 PDF 파일을 업로드해 주세요.")
  else:
    with st.spinner("문서를 분석하고 맞춤법을 검사하는 중입니다..."):
      whitelist, replacements = (
          load_custom_rules(uploaded_excel) if uploaded_excel else (set(), {})
      )
      pdf_bytes = uploaded_pdf.read()
      pdf_lines = extract_text_from_pdf(pdf_bytes)

      results = []
      for item in pdf_lines:
        original_text = item["text"]

        for wrong, right in replacements.items():
          if wrong in original_text:
            results.append({
                "페이지": item["page"],
                "구분": "커스텀 용어",
                "원문": wrong,
                "교정안": right,
                "유형": "프로젝트 용어 통일",
                "사유": "업로드한 기준 엑셀 규칙 적용",
            })

        try:
          spell_result = spell_checker.check(original_text)
          for err in spell_result.errors:
            if err.org_word in whitelist:
              continue
            results.append({
                "페이지": item["page"],
                "구분": "부산대 엔진",
                "원문": err.org_word,
                "교정안": err.cand_word,
                "유형": "맞춤법/띄어쓰기",
                "사유": err.help_text,
            })
        except Exception:
          continue

      if results:
        result_df = pd.DataFrame(results)
        st.success(
            f"검수 완료! 총 {len(result_df)}건의 교정/검토 사항이 발견되었습니다."
        )
        st.dataframe(result_df, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
          result_df.to_excel(writer, index=False, sheet_name="검수결과")
        excel_data = output.getvalue()

        st.download_button(
            label="📥 검수 결과 엑셀 다운로드 (.xlsx)",
            data=excel_data,
            file_name="교정교열_검수결과보고서.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
      else:
        st.info("발견된 오류나 교정 사항이 없습니다.")
