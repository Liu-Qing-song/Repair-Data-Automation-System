import sys
import os
import requests
import json
import time
import re
import tempfile
import uuid
from datetime import datetime
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMessageBox, QProgressDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QScrollArea, QListWidget, QListWidgetItem
from PyQt5.QtCore import QThread, pyqtSignal, QTimer

try:
    from ScreenShots import SingleSnapCapture
    OCR_AVAILABLE = True
except ImportError as e:
    OCR_AVAILABLE = False

class OCRManager:
    _instance = None
    _ocr_capture = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def get_capture(self):
        if self._ocr_capture is None and OCR_AVAILABLE:
            try:
                self._ocr_capture = SingleSnapCapture()
            except Exception as e:
                self._ocr_capture = None
        return self._ocr_capture
    
    def is_available(self):
        return OCR_AVAILABLE and self.get_capture() is not None

class UploadWorker(QThread):
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    record_result = pyqtSignal(str, bool, str)
    finished = pyqtSignal(bool, str, list)
    
    def __init__(self, record_file, task_id, retry_mode=False):
        super().__init__()
        self.record_file = record_file
        self.task_id = task_id
        self.retry_mode = retry_mode
        self.uploader = None
        self.upload_results = []
        self._is_cancelled = False
        
    def cancel(self):
        self._is_cancelled = True
        if self.uploader:
            pass
    
    def categorizeError(self, error_msg):
        if not error_msg:
            return "提交失败"
        
        error_lower = error_msg.lower()
        
        if any(keyword in error_lower for keyword in [
            'connection', 'connect', 'timeout', 'network', '连接', 
            'login', '登录', 'cookie', 'session', 'http'
        ]):
            return "连接失败"
        
        if any(keyword in error_lower for keyword in [
            'search', 'not found', 'no result', '搜索', '未找到', 
            'product', 'fid', 'serial'
        ]):
            return "未查找到产品FID"
        
        if any(keyword in error_lower for keyword in [
            'submit', 'post', 'form', 'data', '提交', '数据'
        ]):
            return "提交失败"
        
        return error_msg[:50] if len(error_msg) > 50 else error_msg
    
    def filterFailedRecords(self, lines):
        filtered_records = []
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            if ' // ' in line:
                original_line, status = line.split(' // ', 1)
                status = status.strip()
                
                if status != 'success':
                    filtered_records.append(original_line.strip())
            else:
                filtered_records.append(line)
        
        return filtered_records
        
    def run(self):
        try:
            if self._is_cancelled:
                return
                
            if not os.path.exists(self.record_file):
                self.finished.emit(False, f"文件不存在: {self.record_file}", [])
                return
            
            self.status.emit("正在连接系统...")
            self.progress.emit(10)
            
            if self._is_cancelled:
                return
            
            try:
                with open(self.record_file, "r", encoding='utf-8') as file:
                    lines = [line.strip() for line in file.readlines() if line.strip()]
            except Exception as e:
                self.finished.emit(False, f"文件读取失败: {str(e)}", [])
                return
            
            if not lines:
                self.finished.emit(False, "文件为空", [])
                return
            
            if self.retry_mode:
                lines = self.filterFailedRecords(lines)
                if not lines:
                    self.finished.emit(True, "🎉 没有失败记录需要重试！", [])
                    return
            
            connection_start = time.time()
            try:
                self.uploader = LowRiskOptimizedUploader()
                self.uploader.checkWebConnection()
                connection_time = time.time() - connection_start
                
                self.status.emit(f"系统连接成功 (耗时: {connection_time:.1f}s)")
                self.progress.emit(20)
            except Exception as e:
                connection_error = self.categorizeError(str(e))
                self.status.emit(f"连接失败: {connection_error}")
                
                failed_results = []
                for i, line in enumerate(lines):
                    if self._is_cancelled:
                        return
                    
                    if self.retry_mode:
                        original_line = line
                    else:
                        if ' // ' in line:
                            original_line = line.split(' // ')[0]
                        else:
                            original_line = line
                    
                    data = [item.strip() for item in original_line.split(',')]
                    product_fid = data[0] if len(data) > 0 else f"记录{i+1}"
                    
                    failed_results.append({
                        "original_line": original_line,
                        "success": False,
                        "error": connection_error,
                        "product_fid": product_fid
                    })
                    
                    self.record_result.emit(product_fid, False, connection_error)
                
                total_records = len(failed_results)
                self.finished.emit(False, f"连接失败: {connection_error}\n❌ 失败：{total_records}条记录", failed_results)
                return
            
            if self._is_cancelled:
                return
            
            self.progress.emit(30)
            success_count = 0
            start_time = time.time()
            
            processed_records = []
            for i, line in enumerate(lines):
                if self._is_cancelled:
                    return
                
                if self.retry_mode:
                    original_line = line
                else:
                    if ' // ' in line:
                        original_line = line.split(' // ')[0]
                    else:
                        original_line = line
                    
                data = [item.strip() for item in original_line.split(',')]
                if len(data) < 13:
                    format_error = "数据格式错误"
                    self.upload_results.append({
                        "original_line": original_line,
                        "success": False,
                        "error": format_error,
                        "product_fid": data[0] if len(data) > 0 else "未知产品"
                    })
                    self.record_result.emit(data[0] if len(data) > 0 else "未知产品", False, format_error)
                    continue
                
                productFID = data[0]
                repairData = {
                    'failureCausedType': data[3], 'repairResult': data[4], 'remarks': data[5],
                    'componentLocation': data[6], 'repairComponentA5E': data[7], 'type': data[8],
                    'failureKind': data[9], 'fcode': data[10], 'repairAction': data[11], 'engineer': data[12]
                }
                processed_records.append((productFID, repairData, original_line))
            
            for i, (productFID, repairData, original_line) in enumerate(processed_records):
                if self._is_cancelled:
                    return
                    
                self.status.emit(f"处理 {i+1}/{len(processed_records)}: {productFID}")
                
                record_start = time.time()
                
                try:
                    result, error_detail = self.uploader.processRepairRecordEnhanced(productFID, repairData)
                    record_time = time.time() - record_start
                    
                    if result:
                        success_count += 1
                        self.upload_results.append({
                            "original_line": original_line,
                            "success": True,
                            "error": "success",
                            "product_fid": productFID
                        })
                        self.record_result.emit(productFID, True, "成功")
                        self.status.emit(f"✅ {productFID} 成功 ({record_time:.1f}s)")
                    else:
                        categorized_error = self.categorizeError(error_detail)
                        self.upload_results.append({
                            "original_line": original_line,
                            "success": False,
                            "error": categorized_error,
                            "product_fid": productFID
                        })
                        self.record_result.emit(productFID, False, categorized_error)
                        self.status.emit(f"❌ {productFID} {categorized_error} ({record_time:.1f}s)")
                        
                except Exception as e:
                    record_time = time.time() - record_start
                    categorized_error = self.categorizeError(str(e))
                    self.upload_results.append({
                        "original_line": original_line,
                        "success": False,
                        "error": categorized_error,
                        "product_fid": productFID
                    })
                    self.record_result.emit(productFID, False, categorized_error)
                    self.status.emit(f"❌ {productFID} {categorized_error} ({record_time:.1f}s)")
                
                progress_value = 30 + int((i + 1) / len(processed_records) * 60)
                self.progress.emit(progress_value)
                
                if self._is_cancelled:
                    return
                
                if success_count > 0:
                    success_rate = success_count / (i + 1)
                    if success_rate > 0.95:
                        time.sleep(0.05)
                    elif success_rate > 0.9:
                        time.sleep(0.08)
                    elif success_rate > 0.7:
                        time.sleep(0.15)
                    else:
                        time.sleep(0.3)
                else:
                    time.sleep(0.15)
            
            if self._is_cancelled:
                return
                
            self.progress.emit(100)
            
            total_time = time.time() - start_time
            total_records = len(self.upload_results)
            failed_count = sum(1 for r in self.upload_results if not r["success"])
            
            mode_prefix = "🔄 重试结果: " if self.retry_mode else ""
            
            if failed_count == 0:
                message = f"{mode_prefix}🎉 全部成功！{success_count}条记录\n⚡ 总耗时:{total_time:.1f}s"
                self.finished.emit(True, message, self.upload_results)
            elif success_count > 0:
                message = f"{mode_prefix}⚠️ 部分成功：{success_count}/{total_records}\n❌ 失败：{failed_count}条\n⚡ 总耗时:{total_time:.1f}s"
                self.finished.emit(True, message, self.upload_results)
            else:
                message = f"{mode_prefix}❌ 全部失败！{failed_count}条记录"
                self.finished.emit(False, message, self.upload_results)
                
        except Exception as e:
            if not self._is_cancelled:
                import traceback
                traceback.print_exc()
                categorized_error = self.categorizeError(str(e))
                self.finished.emit(False, f"上传异常: {categorized_error}", self.upload_results)

class TaskWidget(QtWidgets.QWidget):
    retry_requested = pyqtSignal(str, str)
    record_deleted = pyqtSignal(str, str)
    
    def __init__(self, task_id, filename):
        super().__init__()
        self.task_id = task_id
        self.filename = filename
        self.record_items = {}
        self.setupUI()
        
    def setupUI(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(3)
        
        header_layout = QHBoxLayout()
        self.file_label = QLabel(f"📁 {os.path.basename(self.filename)}")
        self.file_label.setStyleSheet("font-weight: bold; color: #333; font-size: 12px;")
        self.status_label = QLabel("🔄 准备中...")
        self.status_label.setStyleSheet("color: #666; font-size: 11px;")
        
        header_layout.addWidget(self.file_label)
        header_layout.addStretch()
        header_layout.addWidget(self.status_label)
        layout.addLayout(header_layout)
        
        self.progress_bar = QtWidgets.QProgressBar()
        self.progress_bar.setMaximum(100)
        self.progress_bar.setFixedHeight(20)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #ccc;
                border-radius: 5px;
                text-align: center;
                font-size: 11px;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
                border-radius: 4px;
            }
        """)
        layout.addWidget(self.progress_bar)
        
        info_layout = QHBoxLayout()
        self.current_record_label = QLabel("等待开始...")
        self.current_record_label.setStyleSheet("color: #666; font-size: 10px;")
        self.success_count_label = QLabel("✅ 0")
        self.success_count_label.setStyleSheet("color: #4CAF50; font-weight: bold; font-size: 11px;")
        self.fail_count_label = QLabel("❌ 0")
        self.fail_count_label.setStyleSheet("color: #f44336; font-weight: bold; font-size: 11px;")
        
        info_layout.addWidget(self.current_record_label)
        info_layout.addStretch()
        info_layout.addWidget(self.success_count_label)
        info_layout.addWidget(self.fail_count_label)
        layout.addLayout(info_layout)
        
        self.record_details_widget = QListWidget()
        self.record_details_widget.setFixedHeight(120)
        self.record_details_widget.setSelectionMode(QListWidget.NoSelection)
        self.record_details_widget.setStyleSheet("""
            QListWidget {
                border: 1px solid #ddd;
                border-radius: 5px;
                background-color: #fafafa;
                font-size: 10px;
            }
            QListWidget::item {
                padding: 3px 5px;
                border-bottom: 1px solid #eee;
                color: #333;
                min-height: 20px;
            }
        """)
        layout.addWidget(self.record_details_widget)
        
        button_layout = QHBoxLayout()
        
        self.retry_button = QPushButton("🔄 重试")
        self.retry_button.setFixedHeight(30)
        self.retry_button.setStyleSheet("background-color: #FF9800; color: white; font-weight: bold; padding: 3px 10px; font-size: 11px;")
        self.retry_button.clicked.connect(self.requestRetry)
        self.retry_button.setVisible(False)
        
        self.open_file_button = QPushButton("📝 打开文件")
        self.open_file_button.setFixedHeight(30)
        self.open_file_button.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; padding: 3px 10px; font-size: 11px;")
        self.open_file_button.clicked.connect(self.openFile)
        self.open_file_button.setVisible(False)
        
        self.remove_button = QPushButton("🗑️ 移除")
        self.remove_button.setFixedHeight(30)
        self.remove_button.setStyleSheet("background-color: #f44336; color: white; font-weight: bold; padding: 3px 10px; font-size: 11px;")
        self.remove_button.clicked.connect(self.requestRemove)
        self.remove_button.setVisible(False)
        
        button_layout.addStretch()
        button_layout.addWidget(self.retry_button)
        button_layout.addWidget(self.open_file_button)
        button_layout.addWidget(self.remove_button)
        layout.addLayout(button_layout)
        
        self.setFixedHeight(220)
        self.setStyleSheet("""
            TaskWidget {
                background-color: #f9f9f9;
                border: 1px solid #ddd;
                border-radius: 8px;
                margin: 2px;
            }
        """)
        
        self.success_count = 0
        self.fail_count = 0
        
    def updateProgress(self, value):
        self.progress_bar.setValue(value)
        
    def updateStatus(self, status):
        self.status_label.setText(f"🔄 {status}")
        
    def updateCurrentRecord(self, record_info):
        self.current_record_label.setText(f"当前: {record_info}")
        
    def updateRecordResult(self, product_fid, success, error_reason):
        if success:
            self.success_count += 1
            self.success_count_label.setText(f"✅ {self.success_count}")
            self.addRecordDetail(product_fid, True, "成功")
        else:
            self.fail_count += 1
            self.fail_count_label.setText(f"❌ {self.fail_count}")
            self.addRecordDetail(product_fid, False, error_reason)
    
    def addRecordDetail(self, product_fid, success, reason):
        item_widget = QtWidgets.QWidget()
        layout = QHBoxLayout(item_widget)
        layout.setContentsMargins(3, 1, 3, 1)
        layout.setSpacing(5)
        
        if success:
            text_label = QLabel(f"{product_fid}     成功")
            text_label.setStyleSheet("color: #4CAF50; font-weight: bold; font-size: 10px;")
            layout.addWidget(text_label)
            layout.addStretch()
        else:
            display_reason = reason[:15] + "..." if len(reason) > 15 else reason
            text_label = QLabel(f"{product_fid}     {display_reason}")
            text_label.setStyleSheet("color: #f44336; font-weight: bold; font-size: 10px;")
            text_label.setToolTip(f"{product_fid}: {reason}")
            
            delete_button = QPushButton("🗑️")
            delete_button.setFixedSize(18, 18)
            delete_button.setStyleSheet("""
                QPushButton {
                    background-color: #f44336;
                    color: white;
                    border: none;
                    border-radius: 9px;
                    font-size: 9px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #d32f2f;
                }
            """)
            delete_button.clicked.connect(lambda: self.deleteRecord(product_fid))
            
            layout.addWidget(text_label)
            layout.addStretch()
            layout.addWidget(delete_button)
        
        item = QListWidgetItem()
        item.setSizeHint(QtCore.QSize(0, 22))
        self.record_details_widget.addItem(item)
        self.record_details_widget.setItemWidget(item, item_widget)
        
        self.record_items[product_fid] = item
        
        self.record_details_widget.scrollToBottom()
    
    def deleteRecord(self, product_fid):
        try:
            reply = QMessageBox.question(
                self, 
                "确认删除", 
                f"确定要删除记录 {product_fid} 吗？\n\n这将从文件中永久删除该条记录。",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                if product_fid in self.record_items:
                    item = self.record_items[product_fid]
                    row = self.record_details_widget.row(item)
                    self.record_details_widget.takeItem(row)
                    del self.record_items[product_fid]
                    
                    self.fail_count = max(0, self.fail_count - 1)
                    self.fail_count_label.setText(f"❌ {self.fail_count}")
                
                self.record_deleted.emit(self.task_id, product_fid)
                
        except Exception as e:
            QMessageBox.critical(self, "删除失败", f"删除记录时出错:\n{str(e)}")
    
    def updateFilePath(self, new_file_path):
        if new_file_path and os.path.exists(new_file_path):
            self.filename = new_file_path
            self.file_label.setText(f"📁 {os.path.basename(new_file_path)}")
            
    def setCompleted(self, success, message):
        if success:
            if self.fail_count == 0:
                self.status_label.setText("🎉 全部成功")
                self.status_label.setStyleSheet("color: #4CAF50; font-weight: bold; font-size: 11px;")
            else:
                self.status_label.setText("⚠️ 部分成功")
                self.status_label.setStyleSheet("color: #FF9800; font-weight: bold; font-size: 11px;")
        else:
            self.status_label.setText("❌ 失败")
            self.status_label.setStyleSheet("color: #f44336; font-weight: bold; font-size: 11px;")
        
        self.retry_button.setVisible(True)
        self.open_file_button.setVisible(True)
        self.remove_button.setVisible(True)
        self.current_record_label.setText("任务完成")
        
    def requestRetry(self):
        self.retry_requested.emit(self.task_id, self.filename)
    
    def openFile(self):
        try:
            base_name = os.path.splitext(self.filename)[0]
            
            if base_name.endswith('_fail') or base_name.endswith('_done'):
                if base_name.endswith('_fail'):
                    base_name = base_name[:-5]
                elif base_name.endswith('_done'):
                    base_name = base_name[:-5]
            
            possible_files = [
                self.filename,
                f"{base_name}_done.txt",
                f"{base_name}_fail.txt",
                f"{base_name}.txt"
            ]
            
            file_to_open = None
            for file_path in possible_files:
                if os.path.exists(file_path):
                    file_to_open = file_path
                    break
            
            if file_to_open:
                os.startfile(file_to_open)
            else:
                QMessageBox.warning(None, "文件不存在", 
                    f"无法找到文件:\n当前路径: {self.filename}\n尝试的路径:\n" + 
                    "\n".join(f"- {f}" for f in possible_files))
                
        except Exception as e:
            QMessageBox.critical(None, "打开文件失败", f"打开文件时出错:\n{str(e)}")
        
    def requestRemove(self):
        self.setParent(None)
        self.deleteLater()

class TaskManagerWindow(QtWidgets.QDialog):
    retry_task = pyqtSignal(str, str)
    
    def __init__(self):
        super().__init__()
        self.tasks = {}
        self.setupUI()
        
    def setupUI(self):
        self.setWindowTitle("Task Manager")
        self.setGeometry(100, 100, 750, 500)
        
        self.setWindowFlags(
            QtCore.Qt.Window | 
            QtCore.Qt.WindowMinimizeButtonHint | 
            QtCore.Qt.WindowMaximizeButtonHint |
            QtCore.Qt.WindowCloseButtonHint
        )
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(5)
        
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        scroll_area.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        
        self.task_container = QtWidgets.QWidget()
        self.task_layout = QVBoxLayout(self.task_container)
        self.task_layout.setAlignment(QtCore.Qt.AlignTop)
        self.task_layout.setContentsMargins(5, 5, 5, 5)
        self.task_layout.setSpacing(8)
        
        scroll_area.setWidget(self.task_container)
        layout.addWidget(scroll_area)
        
        self.setStyleSheet("""
            QDialog {
                background-color: white;
            }
            QScrollArea {
                border: none;
            }
        """)
        
    def addTask(self, task_id, filename):
        task_widget = TaskWidget(task_id, filename)
        task_widget.retry_requested.connect(self.retry_task.emit)
        
        self.tasks[task_id] = task_widget
        self.task_layout.addWidget(task_widget)
        
        if not self.isVisible():
            self.show()
        self.raise_()
        self.activateWindow()
        
    def removeTask(self, task_id):
        if task_id in self.tasks:
            widget = self.tasks[task_id]
            widget.setParent(None)
            widget.deleteLater()
            del self.tasks[task_id]
    
    def updateTaskFilePath(self, task_id, new_file_path):
        if task_id in self.tasks:
            self.tasks[task_id].updateFilePath(new_file_path)
            
    def updateTaskProgress(self, task_id, progress):
        if task_id in self.tasks:
            self.tasks[task_id].updateProgress(progress)
            
    def updateTaskStatus(self, task_id, status):
        if task_id in self.tasks:
            self.tasks[task_id].updateStatus(status)
            
    def updateTaskRecord(self, task_id, product_fid, success, error_reason):
        if task_id in self.tasks:
            self.tasks[task_id].updateRecordResult(product_fid, success, error_reason)
            
    def setTaskCompleted(self, task_id, success, message):
        if task_id in self.tasks:
            self.tasks[task_id].setCompleted(success, message)
    
    def closeEvent(self, event):
        event.accept()

class TaskManager:
    def __init__(self):
        self.tasks = {}
        self.task_window = TaskManagerWindow()
        self.task_window.retry_task.connect(self.retryTask)
        
    def startNewTask(self, file_path, retry_mode=False):
        task_id = str(uuid.uuid4())[:8]
        
        self.task_window.addTask(task_id, file_path)
        
        if task_id in self.task_window.tasks:
            self.task_window.tasks[task_id].record_deleted.connect(self.deleteRecordFromFile)
        
        worker = UploadWorker(file_path, task_id, retry_mode)
        
        self.tasks[task_id] = {
            'worker': worker,
            'file_path': file_path,
            'original_file': file_path
        }
        
        worker.progress.connect(lambda p: self.task_window.updateTaskProgress(task_id, p))
        worker.status.connect(lambda s: self.task_window.updateTaskStatus(task_id, s))
        worker.record_result.connect(lambda product_fid, s, e: self.task_window.updateTaskRecord(task_id, product_fid, s, e))
        worker.finished.connect(lambda success, msg, results: self.onTaskFinished(task_id, success, msg, results))
        
        worker.start()
        return task_id
        
    def retryTask(self, old_task_id, displayed_file_path):
        try:
            retry_file_path = None
            
            if old_task_id in self.tasks:
                original_file = self.tasks[old_task_id]['original_file']
                
                if os.path.exists(original_file):
                    retry_file_path = original_file
                else:
                    fail_file = original_file.replace('.txt', '_fail.txt')
                    if os.path.exists(fail_file):
                        retry_file_path = fail_file
                    else:
                        done_file = original_file.replace('.txt', '_done.txt')
                        if os.path.exists(done_file):
                            retry_file_path = done_file
            
            if not retry_file_path:
                raise Exception(f"无法找到重试文件: {original_file if old_task_id in self.tasks else displayed_file_path}")
            
            if old_task_id in self.tasks:
                old_worker = self.tasks[old_task_id]['worker']
                
                try:
                    if old_worker and hasattr(old_worker, 'cancel'):
                        old_worker.cancel()
                    
                    if old_worker and hasattr(old_worker, 'isRunning') and old_worker.isRunning():
                        old_worker.terminate()
                        if not old_worker.wait(3000):
                            pass
                            
                except Exception as e:
                    pass
                
                del self.tasks[old_task_id]
            
            self.task_window.removeTask(old_task_id)
            
            new_task_id = self.startNewTask(retry_file_path, retry_mode=True)
            
        except Exception as e:
            QMessageBox.critical(None, "重试失败", f"重试任务时出错:\n{str(e)}")
    
    def deleteRecordFromFile(self, task_id, product_fid):
        try:
            if task_id not in self.tasks:
                return
            
            file_path = self.tasks[task_id]['file_path']
            if not os.path.exists(file_path):
                return
            
            with open(file_path, "r", encoding='utf-8') as file:
                lines = file.readlines()
            
            filtered_lines = []
            deleted_count = 0
            
            for line in lines:
                line_stripped = line.strip()
                if not line_stripped:
                    continue
                
                should_delete = False
                
                if ' // ' in line_stripped:
                    original_part = line_stripped.split(' // ')[0]
                    if self.isExactProductMatch(original_part, product_fid):
                        should_delete = True
                else:
                    if self.isExactProductMatch(line_stripped, product_fid):
                        should_delete = True
                
                if should_delete:
                    deleted_count += 1
                else:
                    filtered_lines.append(line)
            
            if deleted_count > 0:
                with open(file_path, "w", encoding='utf-8') as file:
                    file.writelines(filtered_lines)
                
                QMessageBox.information(None, "删除成功", 
                    f"已从文件中删除记录: {product_fid}\n删除了 {deleted_count} 条相关记录")
            else:
                QMessageBox.warning(None, "删除失败", f"在文件中未找到记录: {product_fid}")
                
        except Exception as e:
            QMessageBox.critical(None, "删除失败", f"删除记录时出错:\n{str(e)}")

    def isExactProductMatch(self, line_content, target_product_fid):
        try:
            parts = [part.strip() for part in line_content.split(',')]
            
            if len(parts) < 1:
                return False
            
            product_fid = parts[0].strip()
            
            return product_fid == target_product_fid
            
        except Exception as e:
            return line_content.startswith(f"{target_product_fid},")
            
    def removeTask(self, task_id):
        try:
            if task_id in self.tasks:
                worker = self.tasks[task_id]['worker']
                
                try:
                    if worker and hasattr(worker, 'cancel'):
                        worker.cancel()
                    
                    if worker and hasattr(worker, 'isRunning') and worker.isRunning():
                        worker.terminate()
                        if not worker.wait(3000):
                            pass
                            
                except Exception as e:
                    pass
                
                del self.tasks[task_id]
                
            self.task_window.removeTask(task_id)
            
        except Exception as e:
            pass
            
    def onTaskFinished(self, task_id, success, message, results):
        try:
            if task_id in self.tasks:
                file_path = self.tasks[task_id]['file_path']
                worker = self.tasks[task_id]['worker']
                
                new_file_path = self.updateFileWithResults(file_path, results)
                
                if new_file_path and new_file_path != file_path:
                    self.tasks[task_id]['file_path'] = new_file_path
                    self.task_window.updateTaskFilePath(task_id, new_file_path)
                
                self.task_window.setTaskCompleted(task_id, success, message)
                
                try:
                    if worker:
                        worker.deleteLater()
                except Exception as e:
                    pass
                
                self.tasks[task_id]['worker'] = None
                
        except Exception as e:
            import traceback
            traceback.print_exc()
            
    def updateFileWithResults(self, original_file, results):
        try:
            if not results:
                return None
                
            new_lines = []
            has_failure = False
            
            for result in results:
                original_line = result["original_line"]
                success = result["success"]
                error = result["error"]
                
                if success:
                    status_text = "success"
                else:
                    has_failure = True
                    status_text = error if error and error != "fail" else "提交失败"
                
                new_line = f"{original_line} // {status_text}\n"
                new_lines.append(new_line)
            
            base_name = os.path.splitext(original_file)[0]
            
            if base_name.endswith('_fail'):
                base_name = base_name[:-5]
            elif base_name.endswith('_done'):
                base_name = base_name[:-5]
            
            suffix = "_fail" if has_failure else "_done"
            new_filename = f"{base_name}{suffix}.txt"
            
            with open(new_filename, "w", encoding='utf-8') as f:
                f.writelines(new_lines)
            
            try:
                if os.path.exists(new_filename) and os.path.exists(original_file) and new_filename != original_file:
                    os.remove(original_file)
            except Exception as e:
                pass
            
            return new_filename
                
        except Exception as e:
            import traceback
            traceback.print_exc()
            return None

class LowRiskOptimizedUploader:
    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'Accept-Language': 'en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7',
            'Accept-Encoding': 'gzip, deflate',
            'Connection': 'keep-alive'
        })
        
        adapter = requests.adapters.HTTPAdapter(
            pool_connections=12,
            pool_maxsize=25,
            max_retries=2
        )
        self.session.mount('http://', adapter)
        self.session.mount('https://', adapter)
        
        self.myCookie = None
        self.requestID = None
        self.uRequestID = None
        
        self.page_cache = {}
        self.search_cache = {}
        self.max_cache_size = 100
        
        self.field_patterns = {
            '__VIEWSTATE': re.compile(r'name="__VIEWSTATE"[^>]*value="([^"]*)"'),
            '__VIEWSTATEGENERATOR': re.compile(r'name="__VIEWSTATEGENERATOR"[^>]*value="([^"]*)"'),
            '__EVENTVALIDATION': re.compile(r'name="__EVENTVALIDATION"[^>]*value="([^"]*)"'),
            
            'txtRequestID': re.compile(r'id="ctl00_ContentPlaceHolder1_txtRequestID"[^>]*value="([^"]*)"'),
            'txtSEWCNoticificaionNo': re.compile(r'id="ctl00_ContentPlaceHolder1_txtSEWCNoticificaionNo"[^>]*value="([^"]*)"'),
            'txtOrderType': re.compile(r'id="ctl00_ContentPlaceHolder1_txtOrderType"[^>]*value="([^"]*)"'),
            'chkisRepeat': re.compile(r'id="ctl00_ContentPlaceHolder1_chkisRepeat"[^>]*checked="checked"'),
            'txtTroubleDesc': re.compile(r'id="ctl00_ContentPlaceHolder1_txtTroubleDesc"[^>]*>([^<]*)</textarea>'),
            
            'cboWorkStationCode': re.compile(r'id="ctl00_ContentPlaceHolder1_cboWorkStationCode"[^>]*>.*?<option[^>]*selected="selected"[^>]*value="([^"]*)"'),
            'txtMLFB': re.compile(r'id="ctl00_ContentPlaceHolder1_txtMLFB"[^>]*value="([^"]*)"'),
            'txtSerialNo': re.compile(r'id="ctl00_ContentPlaceHolder1_txtSerialNo"[^>]*value="([^"]*)"'),
            'txtQty': re.compile(r'id="ctl00_ContentPlaceHolder1_txtQty"[^>]*value="([^"]*)"'),
            'txtUpdatedSerialNo': re.compile(r'id="ctl00_ContentPlaceHolder1_txtUpdatedSerialNo"[^>]*value="([^"]*)"'),
            'chkUpdateSerialNo': re.compile(r'id="ctl00_ContentPlaceHolder1_chkUpdateSerialNo"[^>]*checked="checked"'),
            'txtVSRNumber': re.compile(r'id="ctl00_ContentPlaceHolder1_txtVSRNumber"[^>]*value="([^"]*)"'),
            
            'txtFuntinalStateoriginal': re.compile(r'id="ctl00_ContentPlaceHolder1_txtFuntinalStateoriginal"[^>]*value="([^"]*)"'),
            'txtFuntinalStatelatest': re.compile(r'id="ctl00_ContentPlaceHolder1_txtFuntinalStatelatest"[^>]*value="([^"]*)"'),
            'txtFirmwareoriginal': re.compile(r'id="ctl00_ContentPlaceHolder1_txtFirmwareoriginal"[^>]*value="([^"]*)"'),
            'txtFirmwarelatest': re.compile(r'id="ctl00_ContentPlaceHolder1_txtFirmwarelatest"[^>]*value="([^"]*)"'),
            
            'cboWarranty': re.compile(r'id="ctl00_ContentPlaceHolder1_cboWarranty"[^>]*>.*?<option[^>]*selected="selected"[^>]*value="([^"]*)"'),
            'cboServiceType': re.compile(r'id="ctl00_ContentPlaceHolder1_cboServiceType"[^>]*>.*?<option[^>]*selected="selected"[^>]*value="([^"]*)"'),
            'cboEngineer': re.compile(r'id="ctl00_ContentPlaceHolder1_cboEngineer"[^>]*>.*?<option[^>]*selected="selected"[^>]*value="([^"]*)"'),
            'cboFailureCasedType': re.compile(r'id="ctl00_ContentPlaceHolder1_cboFailureCasedType"[^>]*>.*?<option[^>]*selected="selected"[^>]*value="([^"]*)"'),
            'cboRepairResult': re.compile(r'id="ctl00_ContentPlaceHolder1_cboRepairResult"[^>]*>.*?<option[^>]*selected="selected"[^>]*value="([^"]*)"'),
            
            'dtpConfirmCompleteDate': re.compile(r'id="ctl00_ContentPlaceHolder1_dtpConfirmCompleteDate"[^>]*value="([^"]*)"'),
            'dtpEndRepairDate': re.compile(r'id="ctl00_ContentPlaceHolder1_dtpEndRepairDate"[^>]*value="([^"]*)"'),
            'txtLaborCost': re.compile(r'id="ctl00_ContentPlaceHolder1_txtLaborCost"[^>]*value="([^"]*)"'),
            'chkIsGoodWill': re.compile(r'id="ctl00_ContentPlaceHolder1_chkIsGoodWill"[^>]*checked="checked"'),
            'txtGoodWillNo': re.compile(r'id="ctl00_ContentPlaceHolder1_txtGoodWillNo"[^>]*value="([^"]*)"'),
            
            'txtRemarks': re.compile(r'id="ctl00_ContentPlaceHolder1_txtRemarks"[^>]*>([^<]*)</textarea>'),
            'txtFailureDesc': re.compile(r'id="ctl00_ContentPlaceHolder1_txtFailureDesc"[^>]*>([^<]*)</textarea>'),
            
            'txtPCBA5ENo': re.compile(r'id="ctl00_ContentPlaceHolder1_txtPCBA5ENo"[^>]*value="([^"]*)"'),
            'txtComponentLocation': re.compile(r'id="ctl00_ContentPlaceHolder1_txtComponentLocation"[^>]*value="([^"]*)"'),
            'txtPCBA_FID': re.compile(r'id="ctl00_ContentPlaceHolder1_txtPCBA_FID"[^>]*value="([^"]*)"'),
            'txtRepairedComponentA5E': re.compile(r'id="ctl00_ContentPlaceHolder1_txtRepairedComponentA5E"[^>]*value="([^"]*)"'),
            'cboFailureType': re.compile(r'id="ctl00_ContentPlaceHolder1_cboFailureType"[^>]*>.*?<option[^>]*selected="selected"[^>]*value="([^"]*)"'),
            'txtFCode': re.compile(r'id="ctl00_ContentPlaceHolder1_txtFCode"[^>]*value="([^"]*)"'),
            'cboRepairAction': re.compile(r'id="ctl00_ContentPlaceHolder1_cboRepairAction"[^>]*>.*?<option[^>]*selected="selected"[^>]*value="([^"]*)"'),
            'txtRepairSN': re.compile(r'id="txtRepairSN"[^>]*value="([^"]*)"'),
            'chkBios': re.compile(r'id="chkBios"[^>]*checked="checked"'),
        }
        
        self.common_form_data = {
            'isSubmit': '1',
            'OperationType': 'save'
        }
        
        self.fcode_map = {
            "accu/battery faulty": "F460", "adjustment knob faulty": "F520", "antenna faulty": "F418",
            "ASIC/Gaterray faulty": "F302", "assembly fault": "F295", "Backlight Inverter faulty": "F347",
            "bad solder joint": "F210", "bad via": "F205", "capacitor faulty": "F370",
            "cause of failure not detected (tech./eco.)": "F888", "component missing": "F220",
            "component sheared off": "F221", "connecting terminal not tightened": "F553",
            "connection line interrupted/faulty": "F550", "cover broken": "F505", "diode faulty": "F330",
            "display faulty": "F345", "display wiring faulty": "F348", "EEPROM/FLASH faulty": "F304",
            "electrolytic capacitor faulty": "F371", "EPROM faulty": "F303", "fuse faulty": "F430",
            "heat sink mounting broken": "F515", "housing broken": "F501", "IC faulty": "F300",
            "insulation faulty": "F560", "label faulty": "F235", "LED faulty": "F340",
            "loos contact": "F555", "membrane keyboard faulty": "F446", "Memory Card Slot faulty": "F579",
            "metal chips/whisker": "F272", "microprocessor faulty": "F301",
            "miscellaneous mechanical part missing/damaged": "F590", "nut/screw faulty": "F511",
            "operational amplifier faulty": "F306", "optical fibre faulty": "F580", "optocoupler faulty": "F350",
            "optoMOS-FET relay faulty": "F351", "plug/socket damaged": "F570", "potentiometer faulty": "F365",
            "printed circuit board faulty": "F491", "push button switch faulty": "F445", "quarz faulty": "F390",
            "RAM faulty": "F305", "rectifier faulty": "F332", "relay coil electrically faulty": "F400",
            "resistor faulty": "F360", "screw missing": "F510", "short circuit - connection line": "F551",
            "short circuit - solder bridge": "F212", "short circuit at via": "F206",
            "slider in snap-on-mounting faulty": "F504", "solder joint broken": "F281",
            "switch electrically faulty": "F441", "switch mechanically faulty": "F442", "thyristor faulty": "F326",
            "touch sensor faulty": "F346", "transformer faulty": "F410", "transistor faulty": "F320",
            "triac faulty": "F327", "varistor faulty": "F368", "voltage regulator faulty": "F307",
            "voltage transformer/switching controller faulty": "F308",
            "wrong assembly of component/wrong positioned": "F250", "wrong component": "F230",
            "wrong covering": "F237", "wrong module packaging": "F130", "zener-/suppressor diode faulty": "F331"
        }

    def extractExistingFormData(self, page_content):
        existing_data = {}
        clean_content = page_content.replace('\n', '').replace('\r', '')
        
        for field_name, pattern in self.field_patterns.items():
            try:
                match = pattern.search(clean_content)
                if match:
                    if 'chk' in field_name and 'checked' in pattern.pattern:
                        existing_data[field_name] = True
                    else:
                        existing_data[field_name] = match.group(1).strip()
                else:
                    if 'chk' in field_name:
                        existing_data[field_name] = False
                    else:
                        existing_data[field_name] = ''
            except Exception as e:
                existing_data[field_name] = ''
        
        return existing_data

    def buildCompleteFormData(self, existing_data, repairData, uRequestID):
        failure_kind = repairData.get('failureKind', '')
        fcode = repairData.get('fcode', '')
        if failure_kind and not fcode:
            fcode = self.fcode_map.get(failure_kind, 'F111')
        
        items_list = [
            '', 
            repairData.get('componentLocation', existing_data.get('txtComponentLocation', '')), 
            repairData.get('repairComponentA5E', existing_data.get('txtRepairedComponentA5E', '')),
            repairData.get('type', existing_data.get('cboFailureType', '')), 
            failure_kind, 
            fcode, 
            repairData.get('repairAction', existing_data.get('cboRepairAction', '')),
            '', '0', '', '', '', '', '0'
        ]
        
        form_data = self.common_form_data.copy()
        
        form_data.update({
            '__VIEWSTATE': existing_data.get('__VIEWSTATE', ''),
            '__VIEWSTATEGENERATOR': existing_data.get('__VIEWSTATEGENERATOR', ''),
            '__EVENTVALIDATION': existing_data.get('__EVENTVALIDATION', ''),
        })
        
        form_data.update({
            'RequestID': existing_data.get('txtRequestID', ''),
            'SEWCNoticificaionNo': existing_data.get('txtSEWCNoticificaionNo', ''),
            'OrderType': existing_data.get('txtOrderType', ''),
            'isRepeat': '1' if existing_data.get('chkisRepeat', False) else '0',
            'TroubleDesc': existing_data.get('txtTroubleDesc', ''),
            
            'WorkStationCode': existing_data.get('cboWorkStationCode', ''),
            'MLFB': existing_data.get('txtMLFB', ''),
            'SerialNo': existing_data.get('txtSerialNo', ''),
            'Qty': existing_data.get('txtQty', ''),
            'UpdatedSerialNo': existing_data.get('txtUpdatedSerialNo', ''),
            'UpdateSerialNo': '1' if existing_data.get('chkUpdateSerialNo', False) else '0',
            'VSRNumber': existing_data.get('txtVSRNumber', ''),
            
            'FuntinalStateoriginal': existing_data.get('txtFuntinalStateoriginal', ''),
            'FuntinalStatelatest': existing_data.get('txtFuntinalStatelatest', ''),
            'Firmwareoriginal': existing_data.get('txtFirmwareoriginal', ''),
            'Firmwarelatest': existing_data.get('txtFirmwarelatest', ''),
            
            'Warranty': existing_data.get('cboWarranty', ''),
            'ServiceType': existing_data.get('cboServiceType', ''),
            'ConfirmCompleteDate': existing_data.get('dtpConfirmCompleteDate', ''),
            'EndRepairDate': existing_data.get('dtpEndRepairDate', ''),
            'LaborCost': existing_data.get('txtLaborCost', ''),
            'IsGoodWill': '1' if existing_data.get('chkIsGoodWill', False) else '0',
            'GoodWillNo': existing_data.get('txtGoodWillNo', ''),
            
            'Items': "[" + "$$$".join(items_list) + "]",
            'Remarks': repairData.get('remarks', existing_data.get('txtRemarks', '')),
            'FailureDesc': repairData.get('failureDesc', existing_data.get('txtFailureDesc', '')),
            'RepairResult': repairData.get('repairResult', existing_data.get('cboRepairResult', '')),
            'FailureCasedType': repairData.get('failureCausedType', existing_data.get('cboFailureCasedType', '')),
            'Engineer': repairData.get('engineer', existing_data.get('cboEngineer', '')),
            
            'PCBA5ENo': repairData.get('pcba5eno', existing_data.get('txtPCBA5ENo', '')),
            'ComponentLocation': repairData.get('componentLocation', existing_data.get('txtComponentLocation', '')),
            'PCBA_FID': repairData.get('pcbaFid', existing_data.get('txtPCBA_FID', '')),
            'RepairedComponentA5E': repairData.get('repairComponentA5E', existing_data.get('txtRepairedComponentA5E', '')),
            'FailureType': repairData.get('type', existing_data.get('cboFailureType', '')),
            'FCode': fcode,
            'RepairAction': repairData.get('repairAction', existing_data.get('cboRepairAction', '')),
            'RepairSN': repairData.get('repairSN', existing_data.get('txtRepairSN', '')),
            'Bios': '1' if repairData.get('checkBios', existing_data.get('chkBios', False)) else '0',
            
            'uRequestID': uRequestID
        })
        
        return form_data

    def submitOptimized(self, repairData, pageContent, uRequestID, productFID=""):
        try:
            existing_data = self.extractExistingFormData(pageContent)
            form_data = self.buildCompleteFormData(existing_data, repairData, uRequestID)
            
            response = self.session.post(
                'http://kplus.siemens.com.cn/informationtoolsnew/InterfaceLibrary/SEWC/Repair/RepairOperation.ashx',
                data=form_data, 
                cookies=self.myCookie, 
                headers={
                    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                    'X-Requested-With': 'XMLHttpRequest'
                }, 
                verify=False, 
                timeout=30
            )
            
            return response.status_code == 200
                
        except Exception as e:
            return False

    def checkWebConnection(self):
        try:
            response = self.session.get(
                'http://kplus.siemens.com.cn/informationtoolsnew/', 
                verify=False, 
                timeout=8
            )
            
            if response.status_code != 200:
                raise Exception(f'网站访问失败 (HTTP {response.status_code})')
                
            self.myCookie = response.cookies
            
            payload = {'loginname': 'ting.wen@siemens.com', 'loginpwd': '20150517', 'stype': 'login'}
            login_response = self.session.post(
                'http://kplus.siemens.com.cn/informationtoolsnew/InterfaceLibrary/Login/Login.ashx', 
                data=payload, 
                cookies=self.myCookie, 
                verify=False, 
                timeout=8
            )
            
            if login_response.status_code != 200:
                raise Exception(f'登录请求失败 (HTTP {login_response.status_code})')
                
            self.myCookie = login_response.cookies
            
            login_status = re.findall(r'loginStatus:(\d+),DefaultPage', login_response.text)
            if not login_status or '1' not in login_status:
                raise Exception('登录失败: 用户凭据无效')
            
            system_response = self.session.get(
                'http://kplus.siemens.com.cn/informationtoolsnew/SEWC/Repair/Default.aspx', 
                cookies=self.myCookie, 
                verify=False, 
                timeout=8
            )
            
            if system_response.status_code != 200:
                raise Exception(f'系统访问失败 (HTTP {system_response.status_code})')
            
            return True
            
        except requests.exceptions.ConnectTimeout:
            raise Exception('连接超时: 网络连接缓慢或不稳定')
        except requests.exceptions.ConnectionError:
            raise Exception('连接失败: 无法连接到服务器，请检查网络')
        except requests.exceptions.Timeout:
            raise Exception('请求超时: 服务器无响应')
        except requests.exceptions.RequestException as e:
            raise Exception(f'网络异常: {str(e)}')
        except Exception as e:
            raise e

    def searchProductOptimized(self, productFID):
        if productFID in self.search_cache:
            cached_result = self.search_cache[productFID]
            self.requestID = cached_result['requestID']
            self.uRequestID = cached_result['uRequestID']
            return True
        
        try:
            timeStamp = str(int(time.time() * 1000))
            search_payload = {
                '_search': 'true',
                'nd': timeStamp,
                'rows': '5',
                'page': '1',
                'sidx': 'RequestID',
                'sord': 'desc',
                'filters': json.dumps({
                    "groupOp": "AND",
                    "rules": [{"field": "SerialNo", "op": "eq", "data": productFID}]
                })
            }
            
            response = self.session.post(
                'http://kplus.siemens.com.cn/informationtoolsnew/InterfaceLibrary/SEWC/Repair/Default.ashx',
                data=search_payload, 
                cookies=self.myCookie, 
                headers={'Content-Type': 'application/x-www-form-urlencoded'}, 
                verify=False, 
                timeout=15
            )
            
            if response.status_code == 200:
                result = json.loads(response.text)
                if result.get('records', 0) > 0 and 'rows' in result:
                    for row in result['rows']:
                        if row.get('SerialNo') == productFID:
                            self.requestID = row['RequestID']
                            self.uRequestID = row['uRequestID']
                            
                            if len(self.search_cache) >= self.max_cache_size:
                                oldest_key = next(iter(self.search_cache))
                                del self.search_cache[oldest_key]
                            
                            self.search_cache[productFID] = {
                                'requestID': self.requestID,
                                'uRequestID': self.uRequestID
                            }
                            return True
            return False
                
        except Exception as e:
            return False

    def getEditPageOptimized(self, uRequestID, productFID=""):
        cache_key = f"edit_{uRequestID}"
        if cache_key in self.page_cache:
            return self.page_cache[cache_key]
        
        try:
            edit_url = f'http://kplus.siemens.com.cn/informationtoolsnew/SEWC/Repair/RepairOperation.aspx?sID={uRequestID}'
            response = self.session.get(
                edit_url, 
                verify=False, 
                cookies=self.myCookie, 
                timeout=30
            )
            
            if response.status_code == 200 and 'ctl00$ContentPlaceHolder1$txtRemarks' in response.text:
                if len(self.page_cache) >= self.max_cache_size:
                    oldest_key = next(iter(self.page_cache))
                    del self.page_cache[oldest_key]
                
                self.page_cache[cache_key] = response.text
                return response.text
            return None
            
        except Exception as e:
            return None

    def processRepairRecordEnhanced(self, productFID, repairData):
        try:
            if not self.searchProductOptimized(productFID):
                return False, "未查找到产品FID"
            
            pageContent = self.getEditPageOptimized(self.uRequestID, productFID)
            if not pageContent:
                return False, "无法访问编辑页面"
            
            result = self.submitOptimized(repairData, pageContent, self.uRequestID, productFID)
            
            if result:
                return True, "success"
            else:
                return False, "提交失败"
                
        except Exception as e:
            error_msg = str(e)
            return False, error_msg

    def processRepairRecordOptimized(self, productFID, repairData):
        result, error_detail = self.processRepairRecordEnhanced(productFID, repairData)
        return result

class Ui_Form(object):
    def __init__(self):
        self.labelPass: QtWidgets.QLabel = None
        self.labelFail: QtWidgets.QLabel = None
        self.pushButtonSubmit: QtWidgets.QPushButton = None
        self.pushButtonConfirmFailure: QtWidgets.QPushButton = None
        self.pushButtonClearAll: QtWidgets.QPushButton = None
        self.listWidget: QtWidgets.QListWidget = None
        self.failure_buttons: list = []
        
        self.lineEditProductFID: QtWidgets.QLineEdit = None
        self.lineEditBoardFID1: QtWidgets.QLineEdit = None
        self.lineEditBoardFID2: QtWidgets.QLineEdit = None
        self.lineEditBoardFID3: QtWidgets.QLineEdit = None
        self.lineEditBoardSNR: QtWidgets.QLineEdit = None
        self.lineEditFailureCausedType: QtWidgets.QLineEdit = None
        self.lineEditRemarks: QtWidgets.QLineEdit = None
        self.lineEditComponentLocation: QtWidgets.QLineEdit = None
        self.lineEditRepairComponentA5E: QtWidgets.QLineEdit = None
        self.lineEditFcode: QtWidgets.QLineEdit = None
        
        self.comboBoxRepairResult: QtWidgets.QComboBox = None
        self.comboBoxType: QtWidgets.QComboBox = None
        self.comboBoxFailureKind: QtWidgets.QComboBox = None
        self.comboBoxRepairAction: QtWidgets.QComboBox = None
        self.comboBoxEngineer: QtWidgets.QComboBox = None
        
        self.pushButtonRetryOCR: QtWidgets.QPushButton = None
        self.ocr_auto_triggered = False
        
        self.record_directory: str = ""
        self.current_record_file: str = ""
        self.currentFailureCausedType = None
        self.isFailureTypeLocked = False
        
        self.task_manager = TaskManager()
        
        self.ocr_manager = OCRManager()
        
        self.left_input_sequence = []
        self.right_input_sequence = []
        self.current_left_index = 0
        self.current_right_index = 0
        
        self.failure_kind_data = {
            "0": ["no fault detected"],
            "1": [
                "accu/battery faulty", "adjustment knob faulty", "antenna faulty", "ASIC/Gaterray faulty",
                "assembly fault", "Backlight Inverter faulty", "bad solder joint", "bad via", "capacitor faulty",
                "cause of failure not detected (tech./eco.)", "component missing", "component sheared off",
                "connecting terminal not tightened", "connection line interrupted/faulty", "cover broken",
                "diode faulty", "display faulty", "display wiring faulty", "EEPROM/FLASH faulty",
                "electrolytic capacitor faulty", "EPROM faulty", "fuse faulty", "heat sink mounting broken",
                "housing broken", "IC faulty", "insulation faulty", "label faulty", "LED faulty", "loos contact",
                "membrane keyboard faulty", "Memory Card Slot faulty", "metal chips/whisker", "microprocessor faulty",
                "miscellaneous mechanical part missing/damaged", "nut/screw faulty", "operational amplifier faulty",
                "optical fibre faulty", "optocoupler faulty", "optoMOS-FET relay faulty", "plug/socket damaged",
                "potentiometer faulty", "printed circuit board faulty", "push button switch faulty", "quarz faulty",
                "RAM faulty", "rectifier faulty", "relay coil electrically faulty", "resistor faulty", "screw missing",
                "short circuit - connection line", "short circuit - solder bridge", "short circuit at via",
                "slider in snap-on-mounting faulty", "solder joint broken", "switch electrically faulty",
                "switch mechanically faulty", "thyristor faulty", "touch sensor faulty", "transformer faulty",
                "transistor faulty", "triac faulty", "varistor faulty", "voltage regulator faulty",
                "voltage transformer/switching controller faulty", "wrong assembly of component/wrong positioned",
                "wrong component", "wrong covering", "wrong module packaging", "zener-/suppressor diode faulty"
            ],
            "2": [
                "corrosion - humidity", "corrosion - silver sulfide/corrosive gase", 
                "damaged by shock/drop", "device scratched", "device soiled", 
                "foreign body in device", "manipulation by third party", "overload", 
                "overvoltage", "scorched/melted"
            ],
            "3": [
                "no start up-FW-Update solves problem", "upgrading A-A1 - HW", 
                "upgrading A-A1 - SW"
            ],
            "4": ["transport damage"]
        }
        
        self.complete_fcode_map = {
            "no fault detected": "F000",
            
            "accu/battery faulty": "F460", "adjustment knob faulty": "F520", "antenna faulty": "F418",
            "ASIC/Gaterray faulty": "F302", "assembly fault": "F295", "Backlight Inverter faulty": "F347",
            "bad solder joint": "F210", "bad via": "F205", "capacitor faulty": "F370",
            "cause of failure not detected (tech./eco.)": "F888", "component missing": "F220",
            "component sheared off": "F221", "connecting terminal not tightened": "F553",
            "connection line interrupted/faulty": "F550", "cover broken": "F505", "diode faulty": "F330",
            "display faulty": "F345", "display wiring faulty": "F348", "EEPROM/FLASH faulty": "F304",
            "electrolytic capacitor faulty": "F371", "EPROM faulty": "F303", "fuse faulty": "F430",
            "heat sink mounting broken": "F515", "housing broken": "F501", "IC faulty": "F300",
            "insulation faulty": "F560", "label faulty": "F235", "LED faulty": "F340", "loos contact": "F555",
            "membrane keyboard faulty": "F446", "Memory Card Slot faulty": "F579", "metal chips/whisker": "F272",
            "microprocessor faulty": "F301", "miscellaneous mechanical part missing/damaged": "F590",
            "nut/screw faulty": "F511", "operational amplifier faulty": "F306", "optical fibre faulty": "F580",
            "optocoupler faulty": "F350", "optoMOS-FET relay faulty": "F351", "plug/socket damaged": "F570",
            "potentiometer faulty": "F365", "printed circuit board faulty": "F491", "push button switch faulty": "F445",
            "quarz faulty": "F390", "RAM faulty": "F305", "rectifier faulty": "F332",
            "relay coil electrically faulty": "F400", "resistor faulty": "F360", "screw missing": "F510",
            "short circuit - connection line": "F551", "short circuit - solder bridge": "F212",
            "short circuit at via": "F206", "slider in snap-on-mounting faulty": "F504", "solder joint broken": "F281",
            "switch electrically faulty": "F441", "switch mechanically faulty": "F442", "thyristor faulty": "F326",
            "touch sensor faulty": "F346", "transformer faulty": "F410", "transistor faulty": "F320",
            "triac faulty": "F327", "varistor faulty": "F368", "voltage regulator faulty": "F307",
            "voltage transformer/switching controller faulty": "F308",
            "wrong assembly of component/wrong positioned": "F250", "wrong component": "F230",
            "wrong covering": "F237", "wrong module packaging": "F130", "zener-/suppressor diode faulty": "F331",
            
            "corrosion - humidity": "F930", "corrosion - silver sulfide/corrosive gase": "F939",
            "damaged by shock/drop": "F950", "device scratched": "F992", "device soiled": "F990",
            "foreign body in device": "F995", "manipulation by third party": "F910", "overload": "F921",
            "overvoltage": "F920", "scorched/melted": "F922",
            
            "no start up-FW-Update solves problem": "F688", "upgrading A-A1 - HW": "F710",
            "upgrading A-A1 - SW": "F720",
            
            "transport damage": "X009"
        }

    def updateFailureKindOptions(self, failure_caused_type):
        if failure_caused_type in self.failure_kind_data:
            options = [""] + self.failure_kind_data[failure_caused_type]
            
            self.comboBoxFailureKind.clear()
            self.comboBoxFailureKind.addItems(options)
            
            self.lineEditFcode.clear()
            
            return True
        return False

    def onFailureKindChangedDynamic(self, failure_kind):
        if failure_kind and failure_kind in self.complete_fcode_map:
            fcode = self.complete_fcode_map[failure_kind]
            self.lineEditFcode.setText(fcode)
        elif failure_kind == "":
            self.lineEditFcode.clear()

    def setupInputSequence(self):
        self.left_input_sequence = [
            self.lineEditProductFID,
            self.lineEditBoardFID1,
            self.lineEditBoardFID2,
            self.lineEditBoardFID3,
            self.lineEditBoardSNR
        ]
        
        self.right_input_sequence = [
            self.comboBoxRepairResult,
            self.lineEditRemarks,
            self.lineEditComponentLocation,
            self.lineEditRepairComponentA5E,
            self.comboBoxType,
            self.comboBoxFailureKind,
            self.lineEditFcode,
            self.comboBoxRepairAction,
            self.comboBoxEngineer
        ]
        
        self.current_left_index = 0
        self.current_right_index = 0

    def onLeftEnterPressed(self):
        if self.current_left_index < len(self.left_input_sequence) - 1:
            self.current_left_index += 1
            next_input = self.left_input_sequence[self.current_left_index]
            next_input.setFocus()
            if hasattr(next_input, 'selectAll'):
                next_input.selectAll()
        else:
            self.current_left_index = 0
            first_input = self.left_input_sequence[self.current_left_index]
            first_input.setFocus()
            if hasattr(first_input, 'selectAll'):
                first_input.selectAll()

    def onRightEnterPressed(self):
        if self.current_right_index < len(self.right_input_sequence) - 1:
            self.current_right_index += 1
            next_input = self.right_input_sequence[self.current_right_index]
            next_input.setFocus()
            
            if isinstance(next_input, QtWidgets.QLineEdit):
                next_input.selectAll()
            elif isinstance(next_input, QtWidgets.QComboBox):
                next_input.showPopup()
        else:
            self.current_right_index = 0
            first_input = self.right_input_sequence[self.current_right_index]
            first_input.setFocus()
            
            if isinstance(first_input, QtWidgets.QLineEdit):
                first_input.selectAll()
            elif isinstance(first_input, QtWidgets.QComboBox):
                first_input.showPopup()

    def setupKeyPressEvents(self):
        for i, input_widget in enumerate(self.left_input_sequence):
            original_keyPressEvent = input_widget.keyPressEvent
            
            def createLeftKeyPressEvent(index, original_event):
                def keyPressEvent(event):
                    if event.key() == QtCore.Qt.Key_Return or event.key() == QtCore.Qt.Key_Enter:
                        self.current_left_index = index
                        self.onLeftEnterPressed()
                    else:
                        original_event(event)
                return keyPressEvent
            
            input_widget.keyPressEvent = createLeftKeyPressEvent(i, original_keyPressEvent)

        for i, input_widget in enumerate(self.right_input_sequence):
            original_keyPressEvent = input_widget.keyPressEvent
            
            def createRightKeyPressEvent(index, original_event):
                def keyPressEvent(event):
                    if event.key() == QtCore.Qt.Key_Return or event.key() == QtCore.Qt.Key_Enter:
                        self.current_right_index = index
                        self.onRightEnterPressed()
                    else:
                        original_event(event)
                return keyPressEvent
            
            input_widget.keyPressEvent = createRightKeyPressEvent(i, original_keyPressEvent)

    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(1181, 856)
        font = QtGui.QFont()
        font.setPointSize(11)
        Form.setFont(font)

        self.labelPass = QtWidgets.QLabel(Form)
        self.labelPass.setGeometry(QtCore.QRect(910, 30, 121, 61))
        self.labelPass.setFont(font)
        self.labelPass.setFrameShape(QtWidgets.QFrame.Box)
        self.labelPass.setAlignment(QtCore.Qt.AlignCenter)

        self.labelFail = QtWidgets.QLabel(Form)
        self.labelFail.setGeometry(QtCore.QRect(910, 100, 121, 61))
        self.labelFail.setFont(font)
        self.labelFail.setFrameShape(QtWidgets.QFrame.Box)
        self.labelFail.setAlignment(QtCore.Qt.AlignCenter)

        self.pushButtonConfirmFailure = QtWidgets.QPushButton(Form)
        self.pushButtonConfirmFailure.setGeometry(QtCore.QRect(450, 150, 100, 71))
        self.pushButtonConfirmFailure.setFont(font)
        self.pushButtonConfirmFailure.setText("Confirm")
        self.pushButtonConfirmFailure.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")

        self.pushButtonClearAll = QtWidgets.QPushButton(Form)
        self.pushButtonClearAll.setGeometry(QtCore.QRect(560, 150, 100, 71))
        self.pushButtonClearAll.setFont(font)
        self.pushButtonClearAll.setText("Clear All")
        self.pushButtonClearAll.setStyleSheet("background-color: #f44336; color: white; font-weight: bold;")

        self.pushButtonSubmit = QtWidgets.QPushButton(Form)
        self.pushButtonSubmit.setGeometry(QtCore.QRect(670, 750, 461, 51))
        self.pushButtonSubmit.setFont(font)
        self.pushButtonSubmit.setText("submit all records")

        self.listWidget = QtWidgets.QListWidget(Form)
        self.listWidget.setGeometry(QtCore.QRect(30, 240, 531, 581))
        self.listWidget.setFont(font)

        self.failure_buttons = []
        for i in range(5):
            btn = QtWidgets.QPushButton(Form)
            btn.setGeometry(QtCore.QRect(50 + i * 80, 150, 71, 71))
            btn.setFont(font)
            btn.setText(f"{i}F")
            btn.clicked.connect(lambda checked, x=str(i): self.loadDataForFailureCausedType(x))
            self.failure_buttons.append(btn)

        product_fid_label = QtWidgets.QLabel(Form)
        product_fid_label.setGeometry(QtCore.QRect(30, 50, 81, 31))
        product_fid_label.setFont(font)
        product_fid_label.setAlignment(QtCore.Qt.AlignCenter)
        product_fid_label.setText("产品FID")
        
        self.lineEditProductFID = QtWidgets.QLineEdit(Form)
        self.lineEditProductFID.setGeometry(QtCore.QRect(120, 40, 341, 40))
        self.lineEditProductFID.setFont(font)

        board_fid_label = QtWidgets.QLabel(Form)
        board_fid_label.setGeometry(QtCore.QRect(470, 50, 81, 31))
        board_fid_label.setFont(font)
        board_fid_label.setAlignment(QtCore.Qt.AlignCenter)
        board_fid_label.setText("主板FID")
        
        self.lineEditBoardFID1 = QtWidgets.QLineEdit(Form)
        self.lineEditBoardFID1.setGeometry(QtCore.QRect(560, 40, 105, 40))
        self.lineEditBoardFID1.setFont(font)
        
        self.lineEditBoardFID2 = QtWidgets.QLineEdit(Form)
        self.lineEditBoardFID2.setGeometry(QtCore.QRect(670, 40, 105, 40))
        self.lineEditBoardFID2.setFont(font)
        
        self.lineEditBoardFID3 = QtWidgets.QLineEdit(Form)
        self.lineEditBoardFID3.setGeometry(QtCore.QRect(780, 40, 105, 40))
        self.lineEditBoardFID3.setFont(font)

        snr_label = QtWidgets.QLabel(Form)
        snr_label.setGeometry(QtCore.QRect(470, 100, 81, 31))
        snr_label.setFont(font)
        snr_label.setAlignment(QtCore.Qt.AlignCenter)
        snr_label.setText("SNR")
        
        self.lineEditBoardSNR = QtWidgets.QLineEdit(Form)
        self.lineEditBoardSNR.setGeometry(QtCore.QRect(560, 90, 261, 40))
        self.lineEditBoardSNR.setFont(font)

        self.pushButtonRetryOCR = QtWidgets.QPushButton(Form)
        self.pushButtonRetryOCR.setGeometry(QtCore.QRect(830, 90, 60, 40))
        self.pushButtonRetryOCR.setFont(font)
        self.pushButtonRetryOCR.setText("📷 Retry")
        self.pushButtonRetryOCR.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; font-size: 10px;")
        self.pushButtonRetryOCR.clicked.connect(self.retryOCRCapture)
        
        if not self.ocr_manager.is_available():
            self.pushButtonRetryOCR.setEnabled(False)
            self.pushButtonRetryOCR.setText("❌ OCR")
            self.pushButtonRetryOCR.setStyleSheet("background-color: #999; color: white; font-size: 10px;")

        repair_fields = [
            ("Failure Caused Type", "lineEditFailureCausedType", "input"),
            ("Repair Result", "comboBoxRepairResult", "combo", ["", "Repair ok", "Scrap", "Reject"]),
            ("Remarks", "lineEditRemarks", "input"),
            ("Component Location", "lineEditComponentLocation", "input"),
            ("Repair ComponentA5E", "lineEditRepairComponentA5E", "input"),
            ("Type", "comboBoxType", "combo", ["", "General no defect", "General component or process", "External overstress", "General software or design", "Special case"]),
            ("Failure Kind", "comboBoxFailureKind", "combo", self.getFailureKinds()),
            ("F-Code", "lineEditFcode", "input"),
            ("RepairAction", "comboBoxRepairAction", "combo", ["", "1) Insert", "2) Re-soldering", "3) Re-assembly", "4) Replace", "5) Update SW/HW", "6) Remove", "7) Retest", "8) Scrap", "9) Others"]),
            ("Engineer", "comboBoxEngineer", "combo", ["", "Pan Li", "Gao Yuan", "Duan Wei", "Yang Heng", "Xiong Xiao Ping", "Peng Ying"])
        ]
        
        for i, field_info in enumerate(repair_fields):
            y_pos = 250 + i * 50
            label_text = field_info[0]
            field_name = field_info[1]
            field_type = field_info[2]
            
            label = QtWidgets.QLabel(Form)
            label.setGeometry(QtCore.QRect(600, y_pos, 261, 31))
            label.setFont(font)
            label.setAlignment(QtCore.Qt.AlignCenter)
            label.setText(label_text)
            
            if field_type == "input":
                control = QtWidgets.QLineEdit(Form)
                control.setGeometry(QtCore.QRect(860, y_pos, 291, 40))
                control.setFont(font)
            else:
                control = QtWidgets.QComboBox(Form)
                control.setGeometry(QtCore.QRect(860, y_pos, 291, 40))
                control.setFont(font)
                control.addItems(field_info[3])
            
            setattr(self, field_name, control)
        
        self.comboBoxFailureKind.currentTextChanged.connect(self.onFailureKindChangedDynamic)
        self.pushButtonSubmit.clicked.connect(self.startNewRecord)
        self.pushButtonConfirmFailure.clicked.connect(self.confirmFailureType)
        self.pushButtonClearAll.clicked.connect(self.clearAllData)

        self.lineEditBoardSNR.textChanged.connect(self.onSNRTextChanged)
        self.lineEditBoardFID1.textChanged.connect(self.onBoardFIDChanged)
        self.lineEditBoardFID2.textChanged.connect(self.onBoardFIDChanged)
        self.lineEditBoardFID3.textChanged.connect(self.onBoardFIDChanged)
        
        self.lineEditBoardSNR.focusInEvent = self.onSNRFocusIn

        self.setupInputSequence()
        self.setupKeyPressEvents()

        self.initVariables()
        self.retranslateUi(Form)

    def onSNRFocusIn(self, event):
        QtWidgets.QLineEdit.focusInEvent(self.lineEditBoardSNR, event)
        
        if (self.lineEditBoardSNR.text().strip() == "" and 
            self.ocr_manager.is_available() and 
            not self.ocr_auto_triggered):
            
            self.ocr_auto_triggered = True
            self.performOCRCapture()

    def retryOCRCapture(self):
        if not self.ocr_manager.is_available():
            QMessageBox.warning(None, "OCR不可用", 
                "OCR功能不可用，请检查ScreenShots.py文件是否存在")
            return
        
        self.performOCRCapture()

    def performOCRCapture(self):
        try:
            self.pushButtonRetryOCR.setEnabled(False)
            self.pushButtonRetryOCR.setText("📷 处理中...")
            
            ocr_capture = self.ocr_manager.get_capture()
            if not ocr_capture:
                raise Exception("OCR实例获取失败")
            
            success, v_fields = ocr_capture.capture()
            
            if success and v_fields:
                snr_text = ", ".join(v_fields)
                self.lineEditBoardSNR.setText(snr_text)
                
                self.autoVerifyAndSave()
                
            elif success and not v_fields:
                QMessageBox.information(None, "OCR结果", 
                    "截图识别成功，但未发现V-字段\n\n" +
                    "请确认截图内容包含V-开头的字段\n" +
                    "如需重试，请点击 📷 Retry 按钮")
                
            else:
                QMessageBox.warning(None, "OCR识别失败", 
                    "截图识别失败\n\n" +
                    "可能原因：\n" +
                    "• 未进行截图操作\n" +
                    "• 截图内容无法识别\n" +
                    "• 操作超时（20秒）\n\n" +
                    "如需重试，请点击 📷 Retry 按钮")
                
        except Exception as e:
            QMessageBox.critical(None, "OCR异常", 
                f"OCR截图过程中出错:\n{str(e)}\n\n" +
                "如需重试，请点击 📷 Retry 按钮")
            
        finally:
            self.pushButtonRetryOCR.setEnabled(True)
            self.pushButtonRetryOCR.setText("📷 Retry")

    def onSNRTextChanged(self):
        self.autoVerifyAndSave()

    def onBoardFIDChanged(self):
        self.autoVerifyAndSave()

    def autoVerifyAndSave(self):
        productFID = self.lineEditProductFID.text().strip()
        boardFID1 = self.lineEditBoardFID1.text().strip()
        boardFID2 = self.lineEditBoardFID2.text().strip()
        boardFID3 = self.lineEditBoardFID3.text().strip()
        boardSNR = self.lineEditBoardSNR.text().strip()

        board_fids = [fid for fid in [boardFID1, boardFID2, boardFID3] if fid]
        
        if not board_fids or not boardSNR:
            self.resetPassFailLabels()
            return

        snr_fields = [field.strip() for field in boardSNR.split(',')]
        all_match = all(board_fid in snr_fields for board_fid in board_fids)
        
        if all_match:
            self.labelPass.setText("PASS")
            self.labelPass.setStyleSheet("background-color: green; color: white; font-weight: bold;")
            self.labelFail.setText("FAIL")
            self.labelFail.setStyleSheet("")
            
            if (self.isFailureTypeLocked and 
                productFID and 
                self.currentFailureCausedType is not None):
                
                if not self.current_record_file:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
                    filename = f"repair_batch_{timestamp}.txt"
                    self.current_record_file = os.path.join(self.record_directory, filename)
                
                self.addFIDtoListWidget(productFID, board_fids, self.currentFailureCausedType)
                
                if self.saveToFile(productFID, board_fids, self.currentFailureCausedType):
                    self.clearProductInputsOnly()
                    self.resetPassFailLabels()
                    self.lineEditProductFID.setFocus()
                    self.current_left_index = 0
                else:
                    QMessageBox.critical(None, "保存失败", "数据保存失败，请检查文件权限或磁盘空间")
        else:
            self.labelPass.setText("PASS")
            self.labelPass.setStyleSheet("")
            self.labelFail.setText("FAIL")
            self.labelFail.setStyleSheet("background-color: red; color: white; font-weight: bold;")

    def resetPassFailLabels(self):
        self.labelPass.setText("PASS")
        self.labelPass.setStyleSheet("")
        self.labelFail.setText("FAIL")
        self.labelFail.setStyleSheet("")

    def getFailureKinds(self):
        return [""] + self.failure_kind_data.get("1", [])

    def initVariables(self):
        try:
            self.record_directory = r"C:\Users\z00568pj\Downloads\CsToolUi\CsToolUi\record"
            if not os.path.exists(self.record_directory):
                self.record_directory = os.path.expanduser("~/Documents/RepairTool")
                os.makedirs(self.record_directory, exist_ok=True)
        except:
            self.record_directory = os.getcwd()
        
        self.current_record_file = ""
        self.currentFailureCausedType = None
        self.isFailureTypeLocked = False
        self.ocr_auto_triggered = False

    def onFailureKindChanged(self, failure_kind):
        if self.isFailureTypeLocked:
            return
            
        fcode_map = {
            "accu/battery faulty": "F460", "adjustment knob faulty": "F520", "antenna faulty": "F418",
            "ASIC/Gaterray faulty": "F302", "assembly fault": "F295", "Backlight Inverter faulty": "F347",
            "bad solder joint": "F210", "bad via": "F205", "capacitor faulty": "F370",
            "cause of failure not detected (tech./eco.)": "F888", "component missing": "F220",
            "component sheared off": "F221", "connecting terminal not tightened": "F553",
            "connection line interrupted/faulty": "F550", "cover broken": "F505", "diode faulty": "F330",
            "display faulty": "F345", "display wiring faulty": "F348", "EEPROM/FLASH faulty": "F304",
            "electrolytic capacitor faulty": "F371", "EPROM faulty": "F303", "fuse faulty": "F430",
            "heat sink mounting broken": "F515", "housing broken": "F501", "IC faulty": "F300",
            "insulation faulty": "F560", "label faulty": "F235", "LED faulty": "F340", "loos contact": "F555",
            "membrane keyboard faulty": "F446", "Memory Card Slot faulty": "F579", "metal chips/whisker": "F272",
            "microprocessor faulty": "F301", "miscellaneous mechanical part missing/damaged": "F590",
            "nut/screw faulty": "F511", "operational amplifier faulty": "F306", "optical fibre faulty": "F580",
            "optocoupler faulty": "F350", "optoMOS-FET relay faulty": "F351", "plug/socket damaged": "F570",
            "potentiometer faulty": "F365", "printed circuit board faulty": "F491", "push button switch faulty": "F445",
            "quarz faulty": "F390", "RAM faulty": "F305", "rectifier faulty": "F332",
            "relay coil electrically faulty": "F400", "resistor faulty": "F360", "screw missing": "F510",
            "short circuit - connection line": "F551", "short circuit - solder bridge": "F212",
            "short circuit at via": "F206", "slider in snap-on-mounting faulty": "F504", "solder joint broken": "F281",
            "switch electrically faulty": "F441", "switch mechanically faulty": "F442", "thyristor faulty": "F326",
            "touch sensor faulty": "F346", "transformer faulty": "F410", "transistor faulty": "F320",
            "triac faulty": "F327", "varistor faulty": "F368", "voltage regulator faulty": "F307",
            "voltage transformer/switching controller faulty": "F308",
            "wrong assembly of component/wrong positioned": "F250", "wrong component": "F230",
            "wrong covering": "F237", "wrong module packaging": "F130", "zener-/suppressor diode faulty": "F331"
        }
        
        if failure_kind in fcode_map:
            self.lineEditFcode.setText(fcode_map[failure_kind])
        elif failure_kind == "":
            self.lineEditFcode.clear()

    def loadDataForFailureCausedType(self, failureCausedType):
        if self.isFailureTypeLocked and self.currentFailureCausedType == failureCausedType:
            self.autoVerifyAndSave()
            return
        
        if self.isFailureTypeLocked and self.currentFailureCausedType != failureCausedType:
            self.unlockFailureType()
        
        self.currentFailureCausedType = failureCausedType
        self.lineEditFailureCausedType.setText(failureCausedType)
        
        self.updateFailureKindOptions(failureCausedType)
        
        if failureCausedType == "0":
            self.comboBoxType.setCurrentText("General no defect")
            self.comboBoxFailureKind.setCurrentText("no fault detected")
            self.lineEditFcode.setText("F000")
            self.lineEditRemarks.setText("NA")
            self.lineEditComponentLocation.setText("NA")
            self.lineEditRepairComponentA5E.setText("NA")
            
        elif failureCausedType == "4":
            self.comboBoxType.setCurrentText("Special case")
            self.comboBoxFailureKind.setCurrentText("transport damage")
            self.lineEditFcode.setText("X009")
            self.lineEditRemarks.clear()
            self.lineEditComponentLocation.clear()
            self.lineEditRepairComponentA5E.clear()
            
        else:
            presets = {
                "1": ("General component or process", "", "F111"),
                "2": ("External overstress", "", "F222"),
                "3": ("General software or design", "", "F333")
            }
            
            if failureCausedType in presets:
                type_val, kind_val, fcode_val = presets[failureCausedType]
                self.comboBoxType.setCurrentText(type_val)
                self.comboBoxFailureKind.setCurrentText(kind_val)
                self.lineEditFcode.setText(fcode_val)
                if not self.isFailureTypeLocked:
                    self.lineEditRemarks.clear()
                    self.lineEditComponentLocation.clear()
                    self.lineEditRepairComponentA5E.clear()
        
        if not self.isFailureTypeLocked:
            self.comboBoxRepairResult.setCurrentIndex(0)
            self.comboBoxRepairAction.setCurrentIndex(0)
            self.comboBoxEngineer.setCurrentIndex(0)
        
        self.highlightFailureCausedTypeButton(failureCausedType)

    def confirmFailureType(self):
        if self.currentFailureCausedType is None:
            QMessageBox.warning(None, "选择错误", "请先选择一个Failure类型")
            return
        
        self.isFailureTypeLocked = True
        
        self.pushButtonConfirmFailure.setText("Locked")
        self.pushButtonConfirmFailure.setStyleSheet("background-color: #FF9800; color: white; font-weight: bold;")
        
        for i, btn in enumerate(self.failure_buttons):
            if str(i) == self.currentFailureCausedType:
                btn.setStyleSheet("border: 3px solid #4CAF50; background-color: #E8F5E8; font-weight: bold;")
            else:
                btn.setStyleSheet("border: 1px solid #ccc; background-color: #f0f0f0;")
        
        self.setRightPanelReadOnly(True)
        
        self.autoVerifyAndSave()

    def clearAllData(self):
        self.unlockFailureType()
        
        self.clearAllInputs()
        
        self.resetAllStates()

    def unlockFailureType(self):
        self.isFailureTypeLocked = False
        
        self.pushButtonConfirmFailure.setText("Confirm")
        self.pushButtonConfirmFailure.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        
        self.setRightPanelReadOnly(False)

    def setRightPanelReadOnly(self, readonly):
        input_fields = [
            self.lineEditFailureCausedType, self.lineEditRemarks,
            self.lineEditComponentLocation, self.lineEditRepairComponentA5E, self.lineEditFcode
        ]
        
        for field in input_fields:
            field.setReadOnly(readonly)
            if readonly:
                field.setStyleSheet("background-color: #f0f0f0; color: #666;")
            else:
                field.setStyleSheet("")
        
        combo_boxes = [
            self.comboBoxRepairResult, self.comboBoxType, self.comboBoxFailureKind,
            self.comboBoxRepairAction, self.comboBoxEngineer
        ]
        
        for combo in combo_boxes:
            combo.setEnabled(not readonly)
            if readonly:
                combo.setStyleSheet("background-color: #f0f0f0; color: #666;")
            else:
                combo.setStyleSheet("")

    def highlightFailureCausedTypeButton(self, failureCausedType):
        for i, btn in enumerate(self.failure_buttons):
            if str(i) == failureCausedType:
                if self.isFailureTypeLocked:
                    btn.setStyleSheet("border: 3px solid #4CAF50; background-color: #E8F5E8; font-weight: bold;")
                else:
                    btn.setStyleSheet("border: 2px solid black;")
            else:
                if self.isFailureTypeLocked:
                    btn.setStyleSheet("border: 1px solid #ccc; background-color: #f0f0f0;")
                else:
                    btn.setStyleSheet("")

    def clearFailureCausedTypeSelection(self):
        if not self.isFailureTypeLocked:
            self.currentFailureCausedType = None
            for btn in self.failure_buttons:
                btn.setStyleSheet("")

    def clearProductInputsOnly(self):
        self.lineEditProductFID.clear()
        self.lineEditBoardFID1.clear()
        self.lineEditBoardFID2.clear()
        self.lineEditBoardFID3.clear()
        self.lineEditBoardSNR.clear()
        self.ocr_auto_triggered = False

    def clearAllInputs(self):
        self.clearProductInputsOnly()
        
        for field in [self.lineEditFailureCausedType, self.lineEditRemarks,
                     self.lineEditComponentLocation, self.lineEditRepairComponentA5E, self.lineEditFcode]:
            field.clear()
        for combo in [self.comboBoxRepairResult, self.comboBoxType, self.comboBoxFailureKind, 
                     self.comboBoxRepairAction, self.comboBoxEngineer]:
            combo.setCurrentIndex(0)
        
        self.clearFailureCausedTypeSelection()

    def resetAllStates(self):
        self.resetPassFailLabels()
        self.clearFailureCausedTypeSelection()
        self.currentFailureCausedType = None
        self.isFailureTypeLocked = False
        
        self.pushButtonConfirmFailure.setText("Confirm")
        self.pushButtonConfirmFailure.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        
        self.setRightPanelReadOnly(False)

    def addFIDtoListWidget(self, productFID, board_fids, failureCausedType):
        item_widget = QtWidgets.QWidget()
        h_layout = QtWidgets.QHBoxLayout(item_widget)
        engineer_text = self.comboBoxEngineer.currentText()
        
        board_fids_text = " ".join(board_fids)
        display_text = f"{productFID}, {board_fids_text}, {failureCausedType}F"
        if engineer_text:
            display_text += f", Engineer: {engineer_text}"
        
        label = QtWidgets.QLabel(display_text, item_widget)
        delete_button = QtWidgets.QPushButton("删除", item_widget)
        delete_button.setMaximumWidth(80)
        delete_button.clicked.connect(lambda _, item=item_widget: self.removeItemFromList(item, productFID, board_fids, failureCausedType))
        h_layout.addWidget(label)
        h_layout.addWidget(delete_button)
        h_layout.setContentsMargins(0, 0, 0, 0)
        h_layout.setSpacing(10)
        list_item = QtWidgets.QListWidgetItem(self.listWidget)
        self.listWidget.addItem(list_item)
        self.listWidget.setItemWidget(list_item, item_widget)
        list_item.setSizeHint(item_widget.sizeHint())

    def removeItemFromList(self, item_widget, productFID, board_fids, failureCausedType):
        for i in range(self.listWidget.count()):
            if self.listWidget.itemWidget(self.listWidget.item(i)) == item_widget:
                self.listWidget.takeItem(i)
                break

        if self.current_record_file and os.path.exists(self.current_record_file):
            try:
                with open(self.current_record_file, "r", encoding='utf-8') as file:
                    lines = file.readlines()
                with open(self.current_record_file, "w", encoding='utf-8') as file:
                    for line in lines:
                        board_fids_text = " ".join(board_fids)
                        if not line.startswith(f"{productFID}, {board_fids_text}, {failureCausedType}"):
                            file.write(line)
            except Exception as e:
                pass

    def saveToFile(self, productFID, board_fids, failureCausedType):
        try:
            board_fids_text = " ".join(board_fids)
            line = f"{productFID}, {board_fids_text}, {failureCausedType}, {self.lineEditFailureCausedType.text()}, {self.comboBoxRepairResult.currentText()}, {self.lineEditRemarks.text()}, {self.lineEditComponentLocation.text()}, {self.lineEditRepairComponentA5E.text()}, {self.comboBoxType.currentText()}, {self.comboBoxFailureKind.currentText()}, {self.lineEditFcode.text()}, {self.comboBoxRepairAction.currentText()},{self.comboBoxEngineer.currentText()}\n"

            if not self.current_record_file:
                QMessageBox.warning(None, "文件错误", "记录文件路径未设置。")
                return False

            try:
                directory = os.path.dirname(self.current_record_file)
                if not os.path.exists(directory):
                    os.makedirs(directory, exist_ok=True)
            except Exception as e:
                filename = os.path.basename(self.current_record_file)
                self.current_record_file = os.path.join(os.getcwd(), filename)
            
            success = False
            error_msg = ""
            
            try:
                with open(self.current_record_file, "a", encoding='utf-8', newline='') as file:
                    file.write(line)
                    file.flush()
                success = True
            except Exception as e:
                error_msg = f"标准保存失败: {str(e)}"
            
            if not success:
                try:
                    with open(self.current_record_file, "a", encoding='gbk', newline='') as file:
                        file.write(line)
                        file.flush()
                    success = True
                except Exception as e:
                    error_msg += f"\nGBK编码保存失败: {str(e)}"
            
            if not success:
                try:
                    temp_file = os.path.join(tempfile.gettempdir(), f"repair_backup_{int(time.time())}.txt")
                    with open(temp_file, "a", encoding='utf-8', newline='') as file:
                        file.write(line)
                        file.flush()
                    success = True
                    QMessageBox.information(None, "保存位置变更", f"文件已保存到临时位置:\n{temp_file}")
                    self.current_record_file = temp_file
                except Exception as e:
                    error_msg += f"\n临时文件保存失败: {str(e)}"
            
            if not success:
                QMessageBox.critical(None, "保存失败", f"所有保存方式都失败了:\n{error_msg}")
                return False
            
            return True
            
        except Exception as e:
            error_msg = f"保存过程异常: {str(e)}"
            QMessageBox.critical(None, "保存异常", error_msg)
            return False

    def startNewRecord(self):
        if self.listWidget.count() == 0:
            QMessageBox.warning(None, "提交错误", "列表中需要至少有一个条目")
            return

        if not self.current_record_file or not os.path.exists(self.current_record_file):
            QMessageBox.warning(None, "文件错误", "没有找到要上传的批次文件")
            return
            
        task_id = self.task_manager.startNewTask(self.current_record_file)
        
        self.listWidget.clear()
        self.current_record_file = ""

    def retranslateUi(self, Form):
        Form.setWindowTitle("Repair Tool")
        self.labelPass.setText("PASS")
        self.labelFail.setText("FAIL")
        
        self.lineEditProductFID.setFocus()

if __name__ == "__main__":
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())
