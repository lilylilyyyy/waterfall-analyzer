"""
瀑布图分析工具 - 在线版
基于Streamlit的Web应用，用户上传Excel文件即可生成分析报告
"""

import streamlit as st
import tempfile
import os
import sys
import subprocess

# 页面配置
st.set_page_config(
    page_title="瀑布图分析工具 - 在线版",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义CSS
st.markdown("""
<style>
    .main-header {
        text-align: center;
        color: #1f77b4;
        padding: 20px;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 10px;
        margin-bottom: 30px;
    }
    .main-header h1 {
        color: white;
        margin: 0;
    }
    .upload-section {
        background: #f8f9fa;
        padding: 20px;
        border-radius: 10px;
        margin: 10px 0;
    }
    .success-box {
        background: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
        padding: 15px;
        border-radius: 5px;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

def main():
    """主函数"""

    # 标题
    st.markdown("""
    <div class="main-header">
        <h1>📊 瀑布图分析工具</h1>
        <p style="color: white; margin: 10px 0 0 0;">在线版 v1.0 - 上传文件即可生成专业分析报告</p>
    </div>
    """, unsafe_allow_html=True)

    # 侧边栏说明
    with st.sidebar:
        st.header("📝 使用说明")
        st.markdown("""
        ### 🔢 操作步骤

        **1. 上传文件**
        - 上传"瀑布图数据文件"
        - 上传"工时零件明细文件"

        **2. 生成报告**
        - 点击"生成分析报告"按钮
        - 等待分析完成（约5-10秒）

        **3. 查看和下载**
        - 在线预览报告
        - 下载HTML报告文件

        ---

        ### 📋 文件要求

        **瀑布图数据文件：**
        - 文件名：瀑布图 XLSX 工作表.xlsx
        - 必需列：U列（上月）、V列（当月）
        - 数据区域：B8:V137, AE:AG

        **工时零件明细文件：**
        - 文件名：工时零件明细.xlsx
        - 必需列：A, X, Y, AD, AE, AF, AG, AK, AN, AW
        - 包含完整的项目和商品明细数据

        ---

        ### ✨ 功能特点

        - ✅ 利润环比瀑布图分析
        - ✅ 维修保养TOP20分析
        - ✅ 混合维修TOP20分析
        - ✅ 内部维修TOP10分析
        - ✅ 门店运营成本分析
        - ✅ 智能结论生成

        ---

        ### 💡 提示

        首次使用建议参考示例文件准备数据
        """)

        st.info("🔒 数据安全：上传的文件仅临时存储，处理完成后自动删除")

    # 主要内容区
    st.header("📁 文件上传")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown('<div class="upload-section">', unsafe_allow_html=True)
        st.subheader("📊 瀑布图数据文件")
        waterfall_file = st.file_uploader(
            "请上传 Excel 文件（.xlsx）",
            type=['xlsx'],
            key='waterfall',
            help="包含瀑布图分析所需的环比数据"
        )
        if waterfall_file:
            st.success(f"✓ 已上传: {waterfall_file.name} ({waterfall_file.size / 1024:.1f} KB)")
        st.markdown('</div>', unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="upload-section">', unsafe_allow_html=True)
        st.subheader("📋 工时零件明细文件")
        detail_file = st.file_uploader(
            "请上传 Excel 文件（.xlsx）",
            type=['xlsx'],
            key='detail',
            help="包含项目和商品的详细数据"
        )
        if detail_file:
            st.success(f"✓ 已上传: {detail_file.name} ({detail_file.size / 1024:.1f} KB)")
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("---")

    # 生成报告按钮
    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        generate_btn = st.button(
            "🚀 生成分析报告",
            type="primary",
            use_container_width=True,
            disabled=not (waterfall_file and detail_file)
        )

    if generate_btn:
        if not waterfall_file or not detail_file:
            st.error("❌ 请先上传两个Excel文件！")
            return

        try:
            # 创建临时目录
            with tempfile.TemporaryDirectory() as tmpdir:
                # 保存上传的文件
                waterfall_path = os.path.join(tmpdir, "瀑布图.xlsx")
                detail_path = os.path.join(tmpdir, "明细.xlsx")
                output_path = os.path.join(tmpdir, "报告.html")

                with open(waterfall_path, 'wb') as f:
                    f.write(waterfall_file.getbuffer())
                with open(detail_path, 'wb') as f:
                    f.write(detail_file.getbuffer())

                # 显示进度
                progress_bar = st.progress(0, text="正在初始化...")
                status_text = st.empty()

                # 创建临时配置文件
                config_path = os.path.join(tmpdir, "config.json")
                import json
                config = {
                    "waterfall_excel_path": waterfall_path,
                    "detail_excel_path": detail_path,
                    "output_html_path": output_path
                }
                with open(config_path, 'w', encoding='utf-8') as f:
                    json.dump(config, f, ensure_ascii=False, indent=2)

                progress_bar.progress(20, text="正在读取数据...")

                # 执行分析脚本
                script_path = os.path.join(os.path.dirname(__file__), 'waterfall_analyzer_full.py')

                # 修改脚本执行环境
                import openpyxl
                from datetime import datetime, timedelta
                from collections import defaultdict

                status_text.text("正在分析数据，这可能需要几秒钟...")
                progress_bar.progress(40, text="正在提取基础数据...")

                # 直接执行脚本内容（修改config路径）
                old_dir = os.getcwd()
                os.chdir(tmpdir)

                try:
                    # 复制分析脚本到临时目录
                    import shutil
                    temp_script = os.path.join(tmpdir, "analyzer.py")
                    shutil.copy(script_path, temp_script)

                    progress_bar.progress(60, text="正在生成报告...")

                    # 执行脚本
                    result = subprocess.run(
                        [sys.executable, temp_script],
                        capture_output=True,
                        text=True,
                        timeout=60,
                        cwd=tmpdir
                    )

                    progress_bar.progress(90, text="正在完成...")

                    if result.returncode != 0:
                        st.error(f"❌ 分析出错:\n{result.stderr}")
                        with st.expander("查看详细错误信息"):
                            st.code(result.stderr)
                        return

                finally:
                    os.chdir(old_dir)

                progress_bar.progress(100, text="完成！")
                status_text.empty()
                progress_bar.empty()

                # 读取生成的HTML
                if os.path.exists(output_path):
                    with open(output_path, 'r', encoding='utf-8') as f:
                        html_content = f.read()

                    st.markdown('<div class="success-box">', unsafe_allow_html=True)
                    st.markdown("### ✅ 分析完成！")
                    st.markdown('</div>', unsafe_allow_html=True)

                    # 显示统计信息（从输出中解析）
                    st.subheader("📈 数据统计")
                    if "Data Summary" in result.stdout:
                        summary_lines = result.stdout.split('\n')
                        for line in summary_lines:
                            if "Categories:" in line or "Data Range:" in line:
                                st.text(line.strip())

                    # 创建下载按钮
                    col_dl1, col_dl2, col_dl3 = st.columns([1, 2, 1])
                    with col_dl2:
                        st.download_button(
                            label="📥 下载HTML报告",
                            data=html_content,
                            file_name=f"瀑布图分析报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html",
                            mime="text/html",
                            use_container_width=True
                        )

                    # 在线预览（折叠）
                    with st.expander("👁️ 在线预览报告", expanded=False):
                        st.components.v1.html(html_content, height=800, scrolling=True)

                else:
                    st.error("❌ 未找到生成的报告文件")

        except subprocess.TimeoutExpired:
            st.error("❌ 分析超时，请检查数据文件大小")
        except Exception as e:
            st.error(f"❌ 分析出错: {str(e)}")
            with st.expander("查看详细错误信息"):
                import traceback
                st.code(traceback.format_exc())


if __name__ == "__main__":
    main()
