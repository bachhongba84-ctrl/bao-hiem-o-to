import io
import re
import hmac
import hashlib
import streamlit as st
import pandas as pd
import pytesseract

from PIL import Image, ImageOps, ImageEnhance
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Đọc ảnh bảo hiểm ô tô",
    page_icon="🚗",
    layout="wide"
)

FIELDS = [
    "Phí bảo hiểm mục 7",
    "Tổng phí bảo hiểm mục 8",
    "Chủ xe cơ giới",
    "Địa chỉ",
    "Điện thoại",
    "Số biển kiểm soát",
    "Số khung",
    "Số máy",
    "Năm sản xuất",
    "Loại xe",
    "Trọng tải",
    "Số chỗ ngồi",
    "Mục đích sử dụng",
    "Thời hạn bảo hiểm",
    "Phí bảo hiểm mục 4",
    "Ngày cấp",
    "Số giấy chứng nhận",
]

ALIASES = {
    "Chủ xe cơ giới": ["chủ xe cơ giới"],
    "Địa chỉ": ["địa chỉ"],
    "Điện thoại": ["điện thoại"],
    "Số biển kiểm soát": ["số biển kiểm soát", "biển kiểm soát"],
    "Số khung": ["số khung"],
    "Số máy": ["số máy"],
    "Năm sản xuất": ["năm sx", "năm sản xuất"],
    "Loại xe": ["loại xe"],
    "Trọng tải": ["trọng tải"],
    "Số chỗ ngồi": ["số chỗ ngồi"],
    "Thời hạn bảo hiểm": ["thời hạn bảo hiểm"],
    "Ngày cấp": ["ngày cấp"],
}


def normalize(s):
    return re.sub(r"\s+", " ", s).strip(" .:-\t")


def ocr_image(uploaded):
    image = Image.open(uploaded).convert("RGB")
    image = ImageOps.grayscale(image)
    image = ImageOps.autocontrast(image)
    image = ImageEnhance.Sharpness(image).enhance(2)
    image = image.resize((image.width * 2, image.height * 2))

    text = pytesseract.image_to_string(
        image, lang="vie+eng", config="--psm 11"
    )
    return text


def extract_suggestions(text):
    lines = [normalize(x) for x in text.splitlines()]
    lines = [x for x in lines if x]
    result = {f: "" for f in FIELDS}

    for field, aliases in ALIASES.items():
        for i, line in enumerate(lines):
            low = line.lower()
            found = False
            for alias in aliases:
                pos = low.find(alias)
                if pos >= 0:
                    value = line[pos + len(alias):].strip(" :-.")
                    if value:
                        result[field] = value
                    elif i + 1 < len(lines):
                        result[field] = lines[i + 1]
                    found = True
                    break
            if found:
                break

    # Số chứng nhận thường nằm ở đầu trang 2.
    match = re.search(r"\b\d{2}[A-Z]{1,4}\s*\d{5,15}\b", text, re.I)
    if match:
        result["Số giấy chứng nhận"] = match.group(0)

    # Không suy đoán phí và dấu tích từ các dòng OCR dễ nhầm.
    result["Mục đích sử dụng"] = "Chưa xác định"
    return result


def excel_bytes(records):
    df = pd.DataFrame(records, columns=FIELDS)
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Bao hiem")
        ws = writer.sheets["Bao hiem"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(fill_type="solid", fgColor="1F4E78")
            cell.alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )

        ws.row_dimensions[1].height = 35
        for col in range(1, len(FIELDS) + 1):
            ws.column_dimensions[get_column_letter(col)].width = 24

        guide = writer.book.create_sheet("Huong dan")
        guide.append(["LƯU Ý KIỂM TRA"])
        guide.append(["Đối chiếu tất cả dữ liệu với ảnh gốc."])
        guide.append(["Kiểm tra số giấy, biển số, số khung, số máy."])
        guide.append(["Kiểm tra phí, thời hạn và ngày cấp."])
        guide.append(["Xác nhận dấu tích kinh doanh/không kinh doanh."])
        guide.column_dimensions["A"].width = 75

    return output.getvalue()


def login():
    st.title("🔐 Đăng nhập phần mềm")
    st.caption("Đọc ảnh giấy chứng nhận bảo hiểm ô tô")

    with st.form("login_form"):
        password = st.text_input("Mật khẩu", type="password")
        submitted = st.form_submit_button("Đăng nhập")

    if submitted:
        configured_hash = st.secrets.get("PASSWORD_SHA256", "")
        actual_hash = hashlib.sha256(password.encode("utf-8")).hexdigest()

        if configured_hash and hmac.compare_digest(
            actual_hash, configured_hash.strip().lower()
        ):
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Mật khẩu không đúng hoặc chưa cấu hình Secrets.")


if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    login()
    st.stop()

if "records" not in st.session_state:
    st.session_state["records"] = []

st.title("🚗 Đọc ảnh bảo hiểm ô tô")
st.write("Tải ảnh lên, kiểm tra 17 trường thông tin rồi xuất Excel.")

with st.sidebar:
    st.subheader("Phiên làm việc")
    st.write(f"Đã lưu {len(st.session_state['records'])} bản ghi.")
    if st.button("Đăng xuất"):
        st.session_state["authenticated"] = False
        st.rerun()
    if st.button("Xóa danh sách đang làm"):
        st.session_state["records"] = []
        st.rerun()

uploads = st.file_uploader(
    "Chọn ảnh giấy bảo hiểm (có thể chọn nhiều ảnh)",
    type=["jpg", "jpeg", "png", "tif", "tiff", "bmp"],
    accept_multiple_files=True
)

if uploads:
    names = [f.name for f in uploads]
    selected_name = st.selectbox("Chọn ảnh cần xử lý", names)
    uploaded = next(f for f in uploads if f.name == selected_name)

    left, right = st.columns([1, 1])

    with left:
        st.image(uploaded, caption=uploaded.name, use_container_width=True)
        if st.button("🔎 Đọc chữ trong ảnh"):
            try:
                uploaded.seek(0)
                text = ocr_image(uploaded)
                st.session_state["ocr_text"] = text
                st.session_state["ocr_name"] = uploaded.name
                st.session_state["suggestions"] = extract_suggestions(text)
                st.success("Đã đọc ảnh. Hãy kiểm tra kết quả.")
            except Exception as e:
                st.error(
                    "Không đọc được ảnh. Hãy kiểm tra OCR và "
                    f"dữ liệu tiếng Việt đã cài. Chi tiết: {e}"
                )

    if st.session_state.get("ocr_name") == uploaded.name:
        suggestions = st.session_state["suggestions"]

        with right:
            st.subheader("Kiểm tra và sửa thông tin")
            st.caption(
                "Các trường chưa nhận diện được hãy nhập thủ công. "
                "Không lưu dữ liệu khi chưa kiểm tra."
            )

            with st.form(f"edit_record_{uploaded.name}"):
                values = {}
                for field in FIELDS:
                    # key theo tên ảnh để đổi ảnh thì ô nhập được làm mới
                    key = f"{field}__{uploaded.name}"
                    if field == "Mục đích sử dụng":
                        current = suggestions.get(field, "Chưa xác định")
                        options = [
                            "Chưa xác định",
                            "Kinh doanh",
                            "Không kinh doanh",
                        ]
                        index = options.index(current) if current in options else 0
                        values[field] = st.selectbox(
                            field, options, index=index, key=key
                        )
                    else:
                        values[field] = st.text_input(
                            field, value=suggestions.get(field, ""), key=key
                        )

                save = st.form_submit_button("➕ Thêm giấy này vào danh sách")

            if save:
                values["Tên ảnh nguồn"] = uploaded.name
                st.session_state["records"].append(values)
                st.success("Đã thêm dữ liệu.")

        with st.expander("Xem toàn bộ chữ OCR để đối chiếu"):
            st.text(st.session_state.get("ocr_text", ""))

if st.session_state["records"]:
    st.subheader("Danh sách giấy bảo hiểm đã nhập")
    df = pd.DataFrame(st.session_state["records"])
    st.dataframe(df, use_container_width=True, hide_index=True)

    # Xuất đúng 17 cột theo yêu cầu.
    records_for_export = [
        {field: record.get(field, "") for field in FIELDS}
        for record in st.session_state["records"]
    ]

    st.download_button(
        "📥 Tải file Excel",
        data=excel_bytes(records_for_export),
        file_name="Danh_sach_bao_hiem.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.caption("Dữ liệu OCR có thể sai. Hãy kiểm tra ảnh gốc trước khi xuất file.")
