import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QPushButton, QFileDialog, 
                             QComboBox, QProgressBar, QTextEdit, QGroupBox,
                             QMessageBox, QRadioButton, QButtonGroup, QLineEdit)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon
import time
from dotenv import load_dotenv

# 导入内容重写器
from content_rewriter import ContentRewriter

class WorkerThread(QThread):
    """处理文档的工作线程"""
    update_progress = pyqtSignal(str)
    update_progress_value = pyqtSignal(int, int)  # 添加进度值信号 (当前值, 总值)
    finished_signal = pyqtSignal(str)
    error_signal = pyqtSignal(str)
    
    def __init__(self, file_path, api_type):
        super().__init__()
        self.file_path = file_path
        self.api_type = api_type
        
    def run(self):
        try:
            self.update_progress.emit("初始化内容重写器...")
            rewriter = ContentRewriter(api_type=self.api_type)
            
            # 监听进度回调
            def progress_callback(current, total, message=None):
                self.update_progress_value.emit(current, total)
                if message:
                    self.update_progress.emit(message)
            
            self.update_progress.emit(f"开始处理文件: {os.path.basename(self.file_path)}")
            # 传入进度回调函数
            rewriter.rewrite_content(self.file_path, progress_callback=progress_callback)
            
            self.finished_signal.emit("处理完成！")
        except Exception as e:
            self.error_signal.emit(f"处理时出错: {str(e)}")


class ContentRewriterGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        # 加载环境变量
        load_dotenv()
        
        # 设置窗口
        self.setWindowTitle('内容重写工具')
        self.setGeometry(100, 100, 800, 600)
        
        # 创建主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # 文件选择区域
        file_group = QGroupBox("文件选择")
        file_layout = QHBoxLayout()
        
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setReadOnly(True)
        self.file_path_edit.setPlaceholderText("选择要处理的文档...")
        
        browse_button = QPushButton("浏览...")
        browse_button.clicked.connect(self.browse_file)
        
        file_layout.addWidget(self.file_path_edit, 3)
        file_layout.addWidget(browse_button, 1)
        file_group.setLayout(file_layout)
        main_layout.addWidget(file_group)
        
        # API 设置区域
        api_group = QGroupBox("API 设置")
        api_layout = QVBoxLayout()
        
        # API 类型选择
        api_type_layout = QHBoxLayout()
        api_type_label = QLabel("选择 API 类型:")
        
        self.api_type_group = QButtonGroup(self)
        self.zhipu_radio = QRadioButton("智谱 API")
        self.tongyi_radio = QRadioButton("通义 API")
        self.tongyi_radio.setChecked(True)  # 默认选择通义
        
        self.api_type_group.addButton(self.zhipu_radio, 1)
        self.api_type_group.addButton(self.tongyi_radio, 2)
        
        api_type_layout.addWidget(api_type_label)
        api_type_layout.addWidget(self.zhipu_radio)
        api_type_layout.addWidget(self.tongyi_radio)
        api_type_layout.addStretch()
        
        # API 密钥设置
        key_layout = QHBoxLayout()
        zhipu_key_label = QLabel("智谱 API 密钥:")
        self.zhipu_key_edit = QLineEdit()
        self.zhipu_key_edit.setEchoMode(QLineEdit.Password)
        self.zhipu_key_edit.setText(os.environ.get('ZHIPU_API_KEY', ''))
        
        tongyi_key_label = QLabel("通义 API 密钥:")
        self.tongyi_key_edit = QLineEdit()
        self.tongyi_key_edit.setEchoMode(QLineEdit.Password)
        self.tongyi_key_edit.setText(os.environ.get('TONGYI_API_KEY', ''))
        
        key_layout.addWidget(zhipu_key_label)
        key_layout.addWidget(self.zhipu_key_edit)
        key_layout.addWidget(tongyi_key_label)
        key_layout.addWidget(self.tongyi_key_edit)
        
        api_layout.addLayout(api_type_layout)
        api_layout.addLayout(key_layout)
        api_group.setLayout(api_layout)
        main_layout.addWidget(api_group)
        
        # 进度区域
        progress_group = QGroupBox("处理进度")
        progress_layout = QVBoxLayout()
        
        # 进度条和进度文本
        progress_bar_layout = QHBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)  # 设置为确定进度模式
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        
        self.progress_label = QLabel("0%")
        self.progress_label.setVisible(False)
        
        progress_bar_layout.addWidget(self.progress_bar, 9)
        progress_bar_layout.addWidget(self.progress_label, 1)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        
        progress_layout.addLayout(progress_bar_layout)
        progress_layout.addWidget(self.log_text)
        progress_group.setLayout(progress_layout)
        main_layout.addWidget(progress_group)
        
        # 操作按钮区域
        button_layout = QHBoxLayout()
        
        self.start_button = QPushButton("开始处理")
        self.start_button.clicked.connect(self.start_processing)
        
        open_output_button = QPushButton("打开输出文件夹")
        open_output_button.clicked.connect(self.open_output_folder)
        
        exit_button = QPushButton("退出")
        exit_button.clicked.connect(self.close)
        
        button_layout.addWidget(self.start_button)
        button_layout.addWidget(open_output_button)
        button_layout.addWidget(exit_button)
        
        main_layout.addLayout(button_layout)
        
        # 状态栏
        self.statusBar().showMessage('就绪')
        
        # 初始化日志
        self.log("欢迎使用内容重写工具")
        self.log("请选择要处理的文档并设置API参数")
        
    def browse_file(self):
        """打开文件选择对话框"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择文档", "", "Word文档 (*.docx);;所有文件 (*.*)"
        )
        
        if file_path:
            self.file_path_edit.setText(file_path)
            self.log(f"已选择文档: {file_path}")
    
    def start_processing(self):
        """开始处理文档"""
        file_path = self.file_path_edit.text()
        
        if not file_path:
            QMessageBox.warning(self, "警告", "请先选择要处理的文档！")
            return
        
        # 保存API密钥
        zhipu_key = self.zhipu_key_edit.text()
        tongyi_key = self.tongyi_key_edit.text()
        
        if self.zhipu_radio.isChecked() and not zhipu_key:
            QMessageBox.warning(self, "警告", "请输入智谱 API 密钥！")
            return
        
        if self.tongyi_radio.isChecked() and not tongyi_key:
            QMessageBox.warning(self, "警告", "请输入通义 API 密钥！")
            return
        
        # 设置环境变量
        os.environ['ZHIPU_API_KEY'] = zhipu_key
        os.environ['TONGYI_API_KEY'] = tongyi_key
        
        # 获取API类型
        api_type = "zhipu" if self.zhipu_radio.isChecked() else "tongyi"
        
        # 设置UI状态
        self.start_button.setEnabled(False)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.progress_label.setText("0%")
        self.progress_label.setVisible(True)
        self.statusBar().showMessage('处理中...')
        
        # 创建并启动工作线程
        self.worker = WorkerThread(file_path, api_type)
        self.worker.update_progress.connect(self.log)
        self.worker.update_progress_value.connect(self.update_progress_value)
        self.worker.finished_signal.connect(self.on_process_finished)
        self.worker.error_signal.connect(self.on_process_error)
        self.worker.start()
    
    def update_progress_value(self, current, total):
        """更新进度条值"""
        if total <= 0:
            percentage = 0
        else:
            percentage = int((current / total) * 100)
        
        self.progress_bar.setValue(percentage)
        self.progress_label.setText(f"{percentage}%")
        self.statusBar().showMessage(f'处理中... {current}/{total}')
    
    def on_process_finished(self, message):
        """处理完成回调"""
        self.log(message)
        self.statusBar().showMessage('处理完成')
        self.progress_bar.setValue(100)
        self.progress_label.setText("100%")
        self.progress_bar.setVisible(False)
        self.progress_label.setVisible(False)
        self.start_button.setEnabled(True)
        
        QMessageBox.information(self, "完成", "文档处理完成！")
        
    def on_process_error(self, error_message):
        """处理错误回调"""
        self.log(f"错误: {error_message}")
        self.statusBar().showMessage('处理出错')
        self.progress_bar.setVisible(False)
        self.progress_label.setVisible(False)
        self.start_button.setEnabled(True)
        
        QMessageBox.critical(self, "错误", f"处理文档时出错:\n{error_message}")
    
    def open_output_folder(self):
        """打开输出文件夹"""
        output_dir = os.path.abspath("output")
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        
        # 根据操作系统打开文件夹
        if sys.platform == 'win32':
            os.startfile(output_dir)
        elif sys.platform == 'darwin':  # macOS
            import subprocess
            subprocess.Popen(['open', output_dir])
        else:  # Linux
            import subprocess
            subprocess.Popen(['xdg-open', output_dir])
    
    def log(self, message):
        """添加日志消息"""
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
        # 滚动到底部
        self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # 统一风格
    
    window = ContentRewriterGUI()
    window.show()
    
    sys.exit(app.exec_()) 