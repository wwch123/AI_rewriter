# visualizer.py
import streamlit as st
import os
import io
import tempfile
from PIL import Image
from document_extractor import DocumentExtractor
import json
import re
from functools import lru_cache
import time
import base64
from concurrent.futures import ThreadPoolExecutor
from typing import List, Dict, Any, Optional
from docx import Document
from docx.oxml.ns import qn

# 预编译正则表达式
LATEX_PATTERNS = {
    'inline': re.compile(r'\$[^$]+\$'),
    'display': re.compile(r'\$\$[^$]+\$\$'),
    'env': re.compile(r'\\begin\{.*?\}.*?\\end\{.*?\}'),
    'cmd': re.compile(r'\\[a-zA-Z]+(\{.*?\})*')
}

@st.cache_data
def get_image_base64(image_data: bytes) -> str:
    """将图片转换为base64编码，并缓存结果"""
    try:
        return base64.b64encode(image_data).decode()
    except Exception:
        return None

@st.cache_data
def process_image(image_data: bytes, max_size: tuple = (800, 800)) -> bytes:
    """处理图片大小和格式，并缓存结果"""
    try:
        image = Image.open(io.BytesIO(image_data))
        
        # 调整大小
        if image.width > max_size[0] or image.height > max_size[1]:
            image.thumbnail(max_size, Image.Resampling.LANCZOS)
        
        # 转换格式
        if image.mode not in ('RGB', 'L'):
            image = image.convert('RGB')
        
        # 转换为bytes
        img_byte_arr = io.BytesIO()
        image.save(img_byte_arr, format='JPEG', quality=85, optimize=True)
        return img_byte_arr.getvalue()
    except Exception as e:
        print(f"图片处理错误: {str(e)}")
        return None

@lru_cache(maxsize=1024)
def is_valid_latex(formula: str) -> bool:
    """优化的LaTeX验证"""
    if not isinstance(formula, str) or not formula.strip():
        return False
    return any(pattern.search(formula) for pattern in LATEX_PATTERNS.values())

def create_content_filter():
    """创建内容过滤器"""
    st.sidebar.markdown("### 内容过滤")
    
    # 内容类型过滤
    content_types = st.sidebar.multiselect(
        "选择内容类型",
        ["text", "formula", "image"],
        default=["text", "formula", "image"]
    )
    
    # 文本搜索
    search_text = st.sidebar.text_input("搜索内容", "")
    
    # 页面大小选择
    page_size = st.sidebar.select_slider(
        "每页显示数量",
        options=[5, 10, 15, 20],
        value=10
    )
    
    return content_types, search_text, page_size

def filter_content_blocks(blocks: List[Dict], content_types: List[str], search_text: str) -> List[Dict]:
    """过滤内容块"""
    filtered_blocks = []
    for block in blocks:
        if block['type'] not in content_types:
            continue
        
        if search_text:
            if block['type'] == 'text' and search_text.lower() not in block['content'].lower():
                continue
            if block['type'] == 'formula' and search_text.lower() not in block['content'].lower():
                continue
        
        filtered_blocks.append(block)
    
    return filtered_blocks

def display_structure(structure: List[Dict]):
    """优化的结构显示"""
    if not structure:
        st.info("未检测到文档结构")
        return

    st.subheader("📑 文档结构")
    
    # 创建可折叠的树形结构
    current_level = 0
    for item in structure:
        if item['level'] > current_level:
            st.markdown('<div style="margin-left: {}px">'.format((item['level'] - 1) * 20), unsafe_allow_html=True)
        elif item['level'] < current_level:
            st.markdown('</div>' * (current_level - item['level']), unsafe_allow_html=True)
        
        st.markdown(f"{'#' * item['level']} {item['text']}")
        current_level = item['level']

def display_text_blocks(text_blocks: List[Dict], page_size: int):
    """优化的文本块显示"""
    if not text_blocks:
        st.info("未检测到文本内容")
        return

    # 分页显示
    total_pages = (len(text_blocks) + page_size - 1) // page_size
    if total_pages > 1:
        col1, col2 = st.columns([3, 1])
        with col1:
            page = st.select_slider("选择页面", options=range(1, total_pages + 1), value=1)
        with col2:
            st.markdown(f"**总页数: {total_pages}**")
        
        start_idx = (page - 1) * page_size
        end_idx = min(start_idx + page_size, len(text_blocks))
        current_blocks = text_blocks[start_idx:end_idx]
    else:
        current_blocks = text_blocks

    # 使用容器优化显示
    for idx, block in enumerate(current_blocks, start=1):
        with st.container():
            col1, col2 = st.columns([5, 1])
            with col1:
                st.markdown(f"#### 文本块 {idx}")
                st.write(block['content'])
            with col2:
                if block.get('format_info'):
                    with st.expander("格式"):
                        st.json(block['format_info'])
            st.markdown("---")

def display_images(images: List[Dict], page_size: int):
    """优化的图片显示"""
    if not images:
        st.info("未检测到图片")
        return

    # 预处理图片
    @st.cache_data
    def preprocess_images(imgs):
        processed = []
        for img in imgs:
            processed_data = process_image(img['content'])
            if processed_data:
                processed.append({
                    **img,
                    'processed_content': processed_data,
                    'base64': get_image_base64(processed_data)
                })
        return processed

    processed_images = preprocess_images(images)
    
    # 分页显示
    total_pages = (len(processed_images) + page_size - 1) // page_size
    if total_pages > 1:
        page = st.select_slider("选择图片页面", options=range(1, total_pages + 1), value=1)
        start_idx = (page - 1) * page_size
        end_idx = min(start_idx + page_size, len(processed_images))
        current_images = processed_images[start_idx:end_idx]
    else:
        current_images = processed_images

    # 使用网格布局显示图片
    cols = st.columns(min(3, len(current_images)))
    for idx, (col, img) in enumerate(zip(cols * ((len(current_images) + 2) // 3), current_images)):
        with col:
            if img['base64']:
                st.markdown(
                    f'<img src="data:image/jpeg;base64,{img["base64"]}" style="width:100%">',
                    unsafe_allow_html=True
                )
                st.caption(img.get('description', f'图片 {idx + 1}'))
                with st.expander("详情"):
                    st.json(img.get('position', {}))

def display_block(block: Dict):
    """显示单个内容块"""
    with st.container():
        # 根据内容类型显示不同的内容
        if block['type'] == 'text':
            st.markdown("#### 文本")
            st.write(block['content'])
            
        elif block['type'] == 'image':
            st.markdown("#### 图片")
            col1, col2 = st.columns([2, 1])
            
            with col1:
                # 处理图片
                processed_data = process_image(block['content'])
                if processed_data:
                    base64_img = get_image_base64(processed_data)
                    if base64_img:
                        st.markdown(
                            f'<img src="data:image/jpeg;base64,{base64_img}" style="max-width:100%">',
                            unsafe_allow_html=True
                        )
                    else:
                        st.error("无法显示图片：转换base64失败")
                else:
                    st.error("无法显示图片：处理图片数据失败")
            
            with col2:
                # 显示图片格式和位置信息
                if block.get('format_info'):
                    st.markdown("##### 图片信息")
                    fmt_info = block['format_info']
                    
                    # 显示放置方式
                    st.markdown(f"**放置方式**: {'内嵌' if fmt_info.get('is_inline') else '浮动'}")
                    
                    # 显示尺寸
                    if 'width' in fmt_info and 'height' in fmt_info:
                        width = fmt_info['width']
                        height = fmt_info['height']
                        if width and height:
                            # 转换为厘米
                            width_cm = width / 360000 if width > 1000 else width
                            height_cm = height / 360000 if height > 1000 else height
                            st.markdown(f"**尺寸**: {width_cm:.1f}cm × {height_cm:.1f}cm")
                    
                    # 如果是浮动图片，显示位置信息
                    if not fmt_info.get('is_inline'):
                        st.markdown("##### 位置信息")
                        if 'position_h' in fmt_info:
                            st.markdown(f"**水平参考**: {fmt_info['position_h']}")
                        if 'position_v' in fmt_info:
                            st.markdown(f"**垂直参考**: {fmt_info['position_v']}")
                    
                    # 如果有样式信息
                    if 'style' in fmt_info:
                        with st.expander("样式信息"):
                            st.code(fmt_info['style'])
            
        elif block['type'] == 'heading':
            st.markdown(f"{'#' * (block.get('level', 1) + 1)} {block['content']}")
        
        # 显示格式信息
        if block.get('format_info'):
            with st.expander("格式信息"):
                st.json(block['format_info'])
        
        st.markdown("---")

def display_content_blocks(content_blocks):
    """顺序显示内容块，并清晰展示每个块的边界"""
    if not content_blocks:
        st.warning("没有找到内容块")
        return

    st.subheader("📄 文档内容")
    
    # 使用容器显示所有内容
    with st.container():
        for idx, block in enumerate(content_blocks, 1):
            # 创建一个带边框的容器来显示每个内容块
            with st.container():
                st.markdown(f"### 块 #{idx} ({block['type']})")
                display_block(block)

def get_image_statistics(content_blocks: List[Dict]) -> Dict:
    """获取图片的详细统计信息"""
    stats = {
        'total_images': 0,
        'inline_images': 0,
        'floating_images': 0,
        'shape_images': 0,  # 旧版Word图片
    }
    
    for block in content_blocks:
        if block['type'] == 'image':
            stats['total_images'] += 1
            fmt_info = block.get('format_info', {})
            
            if fmt_info.get('style'):  # 旧版Word图片
                stats['shape_images'] += 1
            elif fmt_info.get('is_inline'):
                stats['inline_images'] += 1
            else:
                stats['floating_images'] += 1
    
    return stats

def display_statistics(content_blocks: List[Dict]):
    """显示文档内容统计信息"""
    # 基本内容类型统计
    content_types = {}
    for block in content_blocks:
        content_types[block['type']] = content_types.get(block['type'], 0) + 1
    
    st.subheader("📊 内容块统计")
    
    # 显示基本统计
    cols = st.columns(len(content_types))
    for col, (type_name, count) in zip(cols, content_types.items()):
        with col:
            if type_name == 'image':
                # 对于图片类型，显示更详细的统计
                image_stats = get_image_statistics(content_blocks)
                st.metric(f"{type_name} 块数", count)
                with st.expander("📸 图片详细统计"):
                    st.markdown(f"""
                    - 内嵌图片: {image_stats['inline_images']}
                    - 浮动图片: {image_stats['floating_images']}
                    - 旧版图片: {image_stats['shape_images']}
                    """)
            else:
                st.metric(f"{type_name} 块数", count)
    
    # 如果有图片，显示详细的图片统计
    if content_types.get('image', 0) > 0:
        st.subheader("📸 图片详细统计")
        image_stats = get_image_statistics(content_blocks)
        
        # 使用列布局显示图片统计
        cols = st.columns(4)
        with cols[0]:
            st.metric("总图片数", image_stats['total_images'])
        with cols[1]:
            st.metric("内嵌图片", image_stats['inline_images'])
        with cols[2]:
            st.metric("浮动图片", image_stats['floating_images'])
        with cols[3]:
            st.metric("旧版图片", image_stats['shape_images'])

def main():
    """主函数"""
    st.set_page_config(
        page_title="文档内容可视化",
        page_icon="📄",
        layout="wide"
    )

    st.title("📄 文档内容可视化")

    # 文件上传
    uploaded_file = st.file_uploader(
        "选择要处理的文档 (支持 .docx)",
        type=['docx']
    )

    if uploaded_file:
        try:
            # 使用临时文件处理上传的文档
            with st.spinner("正在处理文档..."):
                start_time = time.time()
                
                # 创建临时文件
                temp_dir = os.path.join(os.getcwd(), "temp")
                os.makedirs(temp_dir, exist_ok=True)
                temp_file_path = os.path.join(temp_dir, uploaded_file.name)
                
                try:
                    # 保存上传的文件
                    with open(temp_file_path, "wb") as f:
                        f.write(uploaded_file.getvalue())
                    
                    # 处理文档
                    extractor = DocumentExtractor()
                    content = extractor.extract_content(temp_file_path)
                    
                    # 显示统计信息
                    st.success(f"文档处理完成！耗时: {time.time() - start_time:.2f}秒")
                    
                    # 显示统计信息
                    display_statistics(content.get('content_blocks', []))
                    
                    st.markdown("---")

                    # 显示内容
                    display_content_blocks(content.get('content_blocks', []))

                finally:
                    # 清理临时文件
                    if os.path.exists(temp_file_path):
                        os.remove(temp_file_path)
                    
        except Exception as e:
            st.error(f"处理文档时出错: {str(e)}")
            raise

if __name__ == "__main__":
    main()
