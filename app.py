import streamlit as st
import io
import zipfile
import tempfile
import os
import traceback
import sys

st.set_page_config(page_title="PDF 处理平台", layout="wide")

# ================= 状态初始化 =================
if "fragments" not in st.session_state:
    st.session_state.fragments = []
if "step" not in st.session_state:
    st.session_state.step = 1

try:
    import fitz 
    from pdf2docx import Converter
    DEPENDENCIES_LOADED = True
except ImportError as e:
    DEPENDENCIES_LOADED = False
    st.error("环境依赖异常")
    st.warning(f"详细追踪: {e}")
    st.info("请确保安装: pip install PyMuPDF pdf2docx")

# ================= 核心处理逻辑 =================
def universal_to_pdf(file_obj):
    """将传入的各种格式文件统一处理为 PDF 字节流"""
    file_bytes = file_obj.read()
    ext = os.path.splitext(file_obj.name)[1].lower().replace('.', '')
    
    if ext == 'pdf':
        return file_bytes
        
    if ext in ['txt', 'epub', 'xps', 'cbz', 'fb2', 'png', 'jpg', 'jpeg']:
        doc = fitz.open(stream=file_bytes, filetype=ext)
        pdf_bytes = doc.convert_to_pdf()
        doc.close()
        return pdf_bytes
        
    if ext in ['docx']:
        # 核心防护逻辑：检测操作系统
        if sys.platform.startswith('linux'):
            st.error(f"[{file_obj.name}] 解析拦截：云端 Linux 环境暂不支持原生 Word 转换，请先在本地将其另存为 PDF 后再上传。")
            return None
        else:
            try:
                from docx2pdf import convert
                with tempfile.TemporaryDirectory() as temp_dir:
                    docx_path = os.path.join(temp_dir, "temp.docx")
                    pdf_path = os.path.join(temp_dir, "temp.pdf")
                    with open(docx_path, "wb") as f:
                        f.write(file_bytes)
                    convert(docx_path, pdf_path) 
                    with open(pdf_path, "rb") as f:
                        return f.read()
            except Exception as e:
                st.error(f"[{file_obj.name}] 转换失败。请确认本地已安装 MS Word。报错: {e}")
                return None
            
    st.error(f"暂不支持转换该格式: {ext}")
    return None

def parse_page_range(range_str, max_pages):
    pages = set()
    if not range_str or not range_str.strip():
        return list(range(max_pages))
    parts = range_str.split(',')
    for part in parts:
        part = part.strip()
        if '-' in part:
            try:
                start, end = map(int, part.split('-'))
                start, end = max(1, start), min(max_pages, end)
                if start <= end:
                    pages.update(range(start - 1, end))
            except ValueError: pass
        else:
            try:
                page_num = int(part)
                if 1 <= page_num <= max_pages:
                    pages.add(page_num - 1)
            except ValueError: pass
    return sorted(list(pages))

def convert_pdf_to_docx(pdf_bytes):
    with tempfile.TemporaryDirectory() as temp_dir:
        pdf_path, docx_path = os.path.join(temp_dir, "temp.pdf"), os.path.join(temp_dir, "temp.docx")
        with open(pdf_path, "wb") as f: f.write(pdf_bytes)
        cv = Converter(pdf_path)
        cv.convert(docx_path)
        cv.close()
        with open(docx_path, "rb") as f: docx_bytes = f.read()
    return docx_bytes

def get_page_preview(file_bytes, page_index, high_res=True):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    page = doc.load_page(page_index)
    zoom_factor = 4.0 if high_res else 1.0
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom_factor, zoom_factor)) 
    return pix.tobytes("png")

# ================= 交互状态回调 =================
def move_fragment_up(index):
    if index > 0:
        st.session_state.fragments.insert(index - 1, st.session_state.fragments.pop(index))

def move_fragment_down(index):
    if index < len(st.session_state.fragments) - 1:
        st.session_state.fragments.insert(index + 1, st.session_state.fragments.pop(index))

def reset_workflow():
    st.session_state.fragments = []
    st.session_state.step = 1

def sync_page(widget_key, state_key):
    st.session_state[state_key] = st.session_state[widget_key]

# ================= 平台界面 =================
if DEPENDENCIES_LOADED:
    st.title("PDF 文档处理中心")
    
    st.header("01 | 文档导入与范围界定")
    
    uploaded_files = st.file_uploader(
        "拖拽上传文档 (支持 PDF, DOCX, TXT, EPUB, 图片)", 
        accept_multiple_files=True,
        type=['pdf', 'docx', 'txt', 'epub', 'png', 'jpg', 'jpeg'],
        on_change=reset_workflow
    )

    if uploaded_files:
        extraction_configs = {}
        
        for file in uploaded_files:
            with st.expander(f"文档属性: {file.name}", expanded=True):
                try:
                    pdf_bytes = universal_to_pdf(file)
                    if not pdf_bytes:
                        continue 
                        
                    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                    total_pages = doc.page_count
                    
                    state_key = f"page_{file.name}"
                    if state_key not in st.session_state:
                        st.session_state[state_key] = 1
                        
                    slider_key = f"slider_{file.name}"
                    num_key = f"num_{file.name}"
                    
                    col1, col2 = st.columns([1, 2])
                    
                    with col1:
                        nav_col1, nav_col2 = st.columns([2, 1.5])
                        with nav_col1:
                            st.slider("滑动定位", min_value=1, max_value=total_pages, value=st.session_state[state_key], key=slider_key, on_change=sync_page, args=(slider_key, state_key), label_visibility="collapsed")
                        with nav_col2:
                            st.number_input("逐页微调", min_value=1, max_value=total_pages, value=st.session_state[state_key], step=1, key=num_key, on_change=sync_page, args=(num_key, state_key), label_visibility="collapsed")
                        
                        preview_page = st.session_state[state_key]
                        img_bytes = get_page_preview(pdf_bytes, preview_page - 1, high_res=True)
                        st.image(img_bytes, caption=f"Page {preview_page}", use_container_width=True)
                        
                    with col2:
                        st.markdown(f"**文档总页数**: `{total_pages}`")
                        page_range = st.text_input(
                            "提取范围 (例: 1-5, 8)：", 
                            key=f"input_{file.name}",
                            placeholder="留空默认全选"
                        )
                        
                        extraction_configs[file.name] = {
                            "file_bytes": pdf_bytes,
                            "range": page_range,
                            "total_pages": total_pages,
                            "base_name": os.path.splitext(file.name)[0]
                        }
                except Exception as e:
                    st.error(f"解析异常 {file.name}: {e}")

        st.header("02 | 序列编排与输出设定")
        
        col_mode1, col_mode2 = st.columns(2)
        with col_mode1:
            merge_files = st.checkbox("合并为单一文件", value=True)
        with col_mode2:
            output_format = st.selectbox("输出格式", ["PDF", "Word (.docx)"])

        if st.button("构建处理序列", type="primary"):
            st.session_state.fragments = []
            for fname, config in extraction_configs.items():
                parsed_pages = parse_page_range(config["range"], config["total_pages"])
                if parsed_pages:
                    st.session_state.fragments.append({
                        "file_name": fname,
                        "file_bytes": config["file_bytes"],
                        "pages": parsed_pages,
                        "custom_name": f"{config['base_name']}_提取",
                        "first_page_img": get_page_preview(config["file_bytes"], parsed_pages[0], high_res=True) 
                    })
            st.session_state.step = 2

        if st.session_state.step == 2 and st.session_state.fragments:
            st.divider()
            st.subheader("序列视图")
            
            for i, frag in enumerate(st.session_state.fragments):
                frag_col1, frag_col2, frag_col3 = st.columns([1, 3, 1])
                
                with frag_col1:
                    st.image(frag["first_page_img"], caption=f"序列 {i+1}", width=140)
                    
                with frag_col2:
                    st.markdown(f"**源文件**: `{frag['file_name']}`")
                    st.markdown(f"**提取页数**: `{len(frag['pages'])}`")
                    if not merge_files:
                        frag["custom_name"] = st.text_input(
                            f"序列 {i+1} 命名:", 
                            value=frag["custom_name"], 
                            key=f"rename_{i}"
                        )
                        
                with frag_col3:
                    if merge_files:
                        st.markdown("**序列层级**")
                        st.button("上移", key=f"up_{i}", on_click=move_fragment_up, args=(i,), disabled=(i == 0))
                        st.button("下移", key=f"down_{i}", on_click=move_fragment_down, args=(i,), disabled=(i == len(st.session_state.fragments)-1))

            st.divider()
            st.subheader("03 | 任务执行")
            
            if st.button("开始处理", type="primary"):
                with st.spinner("执行中..."):
                    try:
                        if merge_files:
                            merged_doc = fitz.Document()
                            for frag in st.session_state.fragments:
                                src_doc = fitz.open(stream=frag["file_bytes"], filetype="pdf")
                                for page_num in frag["pages"]:
                                    merged_doc.insert_pdf(src_doc, from_page=page_num, to_page=page_num)
                                src_doc.close()
                            
                            out_bytes = merged_doc.write()
                            merged_doc.close()
                            
                            if "Word" in output_format:
                                final_bytes = convert_pdf_to_docx(out_bytes)
                                ext = "docx"
                            else:
                                final_bytes = out_bytes
                                ext = "pdf"
                                
                            st.success("任务完成")
                            st.download_button("📥 获取文件", data=final_bytes, file_name=f"Merged_Document.{ext}", mime="application/octet-stream")
                            
                        else:
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                                for frag in st.session_state.fragments:
                                    src_doc = fitz.open(stream=frag["file_bytes"], filetype="pdf")
                                    new_doc = fitz.Document()
                                    for page_num in frag["pages"]:
                                        new_doc.insert_pdf(src_doc, from_page=page_num, to_page=page_num)
                                    
                                    out_bytes = new_doc.write()
                                    new_doc.close()
                                    src_doc.close()
                                    
                                    safe_name = frag["custom_name"].replace("/", "-").replace("\\", "-") 
                                    if "Word" in output_format:
                                        final_bytes = convert_pdf_to_docx(out_bytes)
                                        ext = "docx"
                                    else:
                                        final_bytes = out_bytes
                                        ext = "pdf"
                                        
                                    zip_file.writestr(f"{safe_name}.{ext}", final_bytes)
                            
                            st.success("任务完成")
                            st.download_button("📥 获取归档压缩包", data=zip_buffer.getvalue(), file_name="Extracted_Documents.zip", mime="application/zip")
                        
                    except Exception as execute_error:
                        st.error("执行中止")
                        st.code(traceback.format_exc())
