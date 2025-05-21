#!/usr/bin/env python
"""
电化学数据处理工具启动脚本
"""
import os
import sys
import traceback
import logging # 新增导入
import tkinter as tk # 添加显示进度窗口
from datetime import datetime # 添加日期时间模块

# 打包为 exe 后，正确获取应用程序目录
def get_application_path():
    """获取应用程序路径，兼容普通运行和 PyInstaller 打包情况"""
    if getattr(sys, 'frozen', False):
        # 如果是通过 PyInstaller 打包的应用程序
        return os.path.dirname(sys.executable)
    else:
        # 如果是普通 Python 脚本
        return os.path.dirname(os.path.abspath(__file__))

# 确保可以找到 electrochemistry 包
# 如果 run_electrochemistry.py 在 'cursor' 目录中，则 'electrochemistry' 包是 script_dir 的子目录
# 并且 'electrochemistry' 也位于 'cursor' 中。
# 因此，script_dir（即 'cursor'）应该在 sys.path 中，以便找到 'electrochemistry'。
# 此外，如果包的通用实用程序或其他部分需要解析路径
# 相对于包的根目录（例如 'electrochemistry' 文件夹本身），
# 则该特定路径可能也需要考虑或通过包内的相对导入来处理。
# 目前，添加 'cursor' (script_dir) 是 'import electrochemistry' 的主要步骤。
script_dir = get_application_path()
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

class ProgressWindow:
    """简单的进度显示窗口，显示程序执行状态"""
    def __init__(self, title="Electrochemical Data Processing Tool"): # 修改默认标题
        self.root = tk.Tk()
        self.root.title(title)

        # 粉色主题颜色
        self.bg_color_light_pink = "#FFF0F5"  # 浅粉色 (薰衣草紫红)
        self.bg_color_medium_pink = "#FFC0CB" # 中粉色 (标准粉色)
        self.text_color_dark_pink = "#C71585" # 深粉色 (中紫红色)
        self.text_color_black = "#333333"    # 深灰色文字
        self.button_color_pink = "#FF69B4"   # 亮粉色 (热粉色)
        self.button_text_color = "white"     # 按钮文字白色

        self.root.configure(bg=self.bg_color_light_pink)

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        width = 550 # 稍微加宽以容纳更美观的布局
        height = 350 # 稍微加高
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        
        # 创建主框架
        frame = tk.Frame(self.root, bg=self.bg_color_light_pink)
        frame.pack(fill="both", expand=True, padx=15, pady=15) # 增加边距
        
        # 创建标题
        title_label = tk.Label(frame, text=title, font=("Arial", 16, "bold"), 
                               fg=self.text_color_dark_pink, bg=self.bg_color_light_pink)
        title_label.pack(pady=(0, 15)) # 调整标题下边距
        
        # 创建日志文本框
        text_frame = tk.Frame(frame, bd=1, relief="sunken") # 给文本框一个边框
        text_frame.pack(fill="both", expand=True)

        self.text = tk.Text(text_frame, height=15, width=60, 
                            bg="#FFFFFF", fg=self.text_color_black, # 白色背景，深灰文字
                            relief="flat", font=("Consolas", 10),
                            padx=5, pady=5) # 使用等宽字体并增加内边距
        
        scrollbar = tk.Scrollbar(text_frame, orient="vertical", command=self.text.yview,
                                 relief="flat", troughcolor=self.bg_color_medium_pink)
        self.text.config(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")
        self.text.pack(side="left", fill="both", expand=True)
        
        # 状态标签
        self.status_var = tk.StringVar()
        self.status_var.set("准备就绪")
        status_label = tk.Label(frame, textvariable=self.status_var, anchor="w",
                                fg=self.text_color_dark_pink, bg=self.bg_color_light_pink,
                                font=("Arial", 9, "italic"))
        status_label.pack(fill="x", pady=(10, 0)) # 调整状态标签上边距
        
        self.root.update()
    
    def log(self, message):
        """添加日志消息到窗口"""
        self.text.insert("end", message + "\n")
        self.text.see("end")  # 自动滚动到底部
        self.root.update()
    
    def set_status(self, message):
        """设置状态栏消息"""
        self.status_var.set(message)
        self.root.update()
    
    def destroy(self):
        """销毁窗口"""
        self.root.destroy()

# 创建自定义日志处理器，将日志输出到进度窗口
class ProgressWindowHandler(logging.Handler):
    """自定义日志处理器，将日志输出到进度窗口"""
    def __init__(self, window):
        logging.Handler.__init__(self)
        self.window = window
    
    def emit(self, record):
        log_entry = self.format(record)
        self.window.log(log_entry)

def main_entry(): # 从 main 重命名而来
    # 创建进度窗口
    progress_window = ProgressWindow("Electrochemical Data Processing Tool") # 修改实例化时的标题
    progress_window.log("应用程序正在启动...")
    progress_window.set_status("正在初始化...")
    
    # 在此设置日志记录，在执行任何其他操作之前，以确保尽早配置。
    # 这是一个基本配置。electrochemistry.main.setup_environment
    # 之后可以根据需要添加文件处理器或更具体的配置。
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(module)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # 获取此模块的日志记录器
    logger = logging.getLogger(__name__)
    
    # 添加进度窗口日志处理器
    window_handler = ProgressWindowHandler(progress_window)
    window_handler.setFormatter(logging.Formatter('%(levelname)s: %(message)s'))
    logging.getLogger().addHandler(window_handler)
    
    progress_window.log("日志系统已初始化。")

    try:
        # 设置日志文件目录
        app_path = get_application_path()
        log_dir = os.path.join(app_path, "logs")
        os.makedirs(log_dir, exist_ok=True)
        
        # 添加文件处理器
        log_file = os.path.join(log_dir, f"electrochemistry_{datetime.now().strftime('%Y%m%d')}.log")
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(module)s - %(message)s'))
        logging.getLogger().addHandler(file_handler)
        
        progress_window.log(f"日志将保存到: {log_file}")
        progress_window.set_status("正在加载模块...")
        
        logger.info("run_electrochemistry.py: 启动脚本开始执行")
        
        try:
            from electrochemistry.main import main as process_electrochemistry_data
            progress_window.log("主模块已成功加载。")
            progress_window.set_status("主程序启动中...")
            
            logger.info("run_electrochemistry.py: electrochemistry.main 模块已导入")
            process_electrochemistry_data()
            logger.info("run_electrochemistry.py: process_electrochemistry_data 函数已执行")
            
        except ImportError as e:
            error_msg = f"错误：无法导入电化学数据处理包: {e}"
            logger.error(error_msg, exc_info=True)
            progress_window.log(error_msg)
            progress_window.set_status("导入错误")
            
            # 添加更详细的错误信息
            if hasattr(e, 'name') and (e.name == 'electrochemistry.main' or e.name == 'main' or e.name == 'electrochemistry'):
                detail_msg = f"看起来 '{e.name}' 模块或包无法找到。当前 sys.path: {sys.path}. 期望 'electrochemistry' 包在以下目录中: {script_dir}"
                logger.error(detail_msg)
                progress_window.log(detail_msg)
            elif hasattr(e, 'name') and e.name:
                detail_msg = f"无法导入名为 '{e.name}' 的模块或包。"
                logger.error(detail_msg)
                progress_window.log(detail_msg)
            else:
                detail_msg = "发生了一个未指明名称的导入错误。"
                logger.error(detail_msg)
                progress_window.log(detail_msg)
                
    except KeyboardInterrupt:
        logger.warning("程序被用户中断")
        progress_window.log("程序被用户中断")
        progress_window.set_status("已中断")
        
    except Exception as e:
        error_msg = f"运行过程中发生严重错误: {e}"
        logger.critical(error_msg, exc_info=True)
        progress_window.log(error_msg)
        progress_window.log(traceback.format_exc())
        progress_window.set_status("发生错误")
        
    finally:
        logger.info("run_electrochemistry.py: 启动脚本执行完毕或已终止")
        progress_window.log("处理完成。")
        progress_window.set_status("已完成！")
        
        # 添加一个按钮用于关闭程序
        close_button = tk.Button(progress_window.root, text="关闭程序", 
                                 command=progress_window.root.destroy,
                                 bg=progress_window.button_color_pink, 
                                 fg=progress_window.button_text_color,
                                 font=("Arial", 10, "bold"),
                                 relief="raised", padx=10, pady=5)
        close_button.pack(pady=(10, 15)) # 调整关闭按钮边距
        
        # 等待用户关闭窗口
        progress_window.root.mainloop()

if __name__ == "__main__":
    main_entry() # 调用重命名后的 main 函数

