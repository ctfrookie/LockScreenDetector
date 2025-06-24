import win32serviceutil
import win32service
import win32event
import servicemanager
import win32gui
import win32con
import pythoncom
import time
import requests
import json
import os
import sys
import logging
import wmi

# 配置文件路径
CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json')

# 默认配置
DEFAULT_CONFIG = {
    "webhook_url": "",
    "check_interval": 2,
    "log_level": "INFO",
    "detection_method": "wmi"  # 检测方法: wmi 或 polling
}

class ConfigManager:
    """配置管理器"""
    def __init__(self):
        self.config = DEFAULT_CONFIG
        self.load_config()
    
    def load_config(self):
        """加载配置文件"""
        try:
            if os.path.exists(CONFIG_PATH):
                with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
                    # 合并默认配置，防止缺少字段
                    self.config = {**DEFAULT_CONFIG, **self.config}
            else:
                self.save_config()  # 创建默认配置文件
        except Exception as e:
            print(f"加载配置失败: {e}")
    
    def save_config(self):
        """保存配置文件"""
        try:
            with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"保存配置失败: {e}")
    
    def get(self, key, default=None):
        """获取配置项"""
        return self.config.get(key, default)
    
    def set(self, key, value):
        """设置配置项"""
        self.config[key] = value
        self.save_config()

class LockScreenService(win32serviceutil.ServiceFramework):
    """Windows锁屏检测服务"""
    _svc_name_ = "LockScreenDetector"
    _svc_display_name_ = "锁屏检测服务"
    _svc_description_ = "检测Windows系统锁屏状态并发送通知"
    
    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.config = ConfigManager()
        self.last_locked = False
        self.logger = self.setup_logger()
        self.is_running = False
        self.wmi_obj = None
        
        # 记录服务初始化
        self.logger.info("服务初始化完成")
    
    def setup_logger(self):
        """设置日志记录器"""
        logger = logging.getLogger('LockScreenService')
        
        # 设置日志级别
        log_level = self.config.get("log_level", "INFO").upper()
        if log_level == "DEBUG":
            logger.setLevel(logging.DEBUG)
        elif log_level == "WARNING":
            logger.setLevel(logging.WARNING)
        elif log_level == "ERROR":
            logger.setLevel(logging.ERROR)
        else:
            logger.setLevel(logging.INFO)
        
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        
        # 确保日志目录存在
        log_dir = os.path.dirname(os.path.abspath(__file__))
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        # 添加文件处理器
        log_file = os.path.join(log_dir, 'service.log')
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setFormatter(formatter)
        
        # 清除现有处理器，防止重复日志
        if logger.handlers:
            for handler in logger.handlers:
                logger.removeHandler(handler)
                
        logger.addHandler(file_handler)
        
        # 添加控制台处理器（仅在调试模式下）
        if log_level == "DEBUG":
            console_handler = logging.StreamHandler()
            console_handler.setFormatter(formatter)
            logger.addHandler(console_handler)
        
        return logger
    
    def SvcStop(self):
        """停止服务"""
        self.logger.info("收到服务停止请求")
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.is_running = False
        win32event.SetEvent(self.hWaitStop)
        self.logger.info("服务已标记为停止")
    
    def SvcDoRun(self):
        """运行服务"""
        try:
            self.logger.info("服务开始运行")
            self.ReportServiceStatus(win32service.SERVICE_RUNNING)
            self.is_running = True
            
            # 初始化WMI对象
            try:
                self.wmi_obj = wmi.WMI()
                self.logger.info("WMI服务初始化成功")
            except Exception as e:
                self.logger.error(f"WMI服务初始化失败: {e}")
            
            # 选择检测方法
            detection_method = self.config.get("detection_method", "wmi")
            if detection_method.lower() == "wmi":
                self.logger.info("使用WMI方法检测锁屏状态")
                self.run_with_wmi()
            else:
                self.logger.info("使用轮询方法检测锁屏状态")
                self.main()
        except Exception as e:
            # 捕获致命异常，记录并报告服务失败
            self.logger.critical(f"服务运行时发生致命错误: {str(e)}", exc_info=True)
            self.ReportServiceStatus(win32service.SERVICE_STOPPED, win32service.SERVICE_ERROR_CRITICAL)
    
    def is_screen_locked_wmi(self):
        """使用WMI检测屏幕是否锁定"""
        if not self.wmi_obj:
            return False, None
            
        try:
            # 获取当前登录用户
            current_user = None
            for session in self.wmi_obj.Win32_ComputerSystem():
                current_user = session.UserName
                self.logger.debug(f"当前登录用户: {current_user}")
            
            # 如果没有用户登录，可能是锁屏状态
            if not current_user:
                self.logger.debug("没有用户登录，判定为锁屏状态")
                return True, current_user
                
            # 获取工作站锁定状态
            for session in self.wmi_obj.Win32_Process(name="LogonUI.exe"):
                self.logger.debug("找到LogonUI进程，判定为锁屏状态")
                return True, current_user
                
            self.logger.debug("未检测到锁屏状态")
            return False, current_user
        except Exception as e:
            self.logger.error(f"WMI检测锁屏状态失败: {e}", exc_info=True)
            # 出错时默认返回False和None，避免误报锁屏状态
            return False, None
    
    def lock_state_changed(self, is_locked, user):
        """处理锁屏状态变化"""
        try:
            user_display = user if user else "未知用户"
            
            if is_locked and not self.last_locked:
                message = f"用户 {user_display} 已锁定系统"
                self.logger.info(message)
                self.send_wechat_notification(message)
            elif not is_locked and self.last_locked:
                message = f"用户 {user_display} 已解锁系统"
                self.logger.info(message)
                self.send_wechat_notification(message)
                
            self.last_locked = is_locked
        except Exception as e:
            self.logger.error(f"处理锁屏状态变化失败: {e}", exc_info=True)
    
    def send_wechat_notification(self, message):
        """发送微信通知"""
        webhook_url = self.config.get("webhook_url")
        if not webhook_url:
            self.logger.warning("未配置Webhook URL，无法发送通知")
            return
        
        try:
            self.logger.info(f"正在发送微信通知: {message}")
            response = requests.post(webhook_url, json={
                "msgtype": "text",
                "text": {
                    "content": message
                }
            }, timeout=10)
            
            if response.status_code == 200:
                result = response.json()
                if result.get("errcode") == 0:
                    self.logger.info("微信通知发送成功")
                else:
                    self.logger.error(f"微信API返回错误: {result}")
            else:
                self.logger.error(f"微信通知发送失败，状态码: {response.status_code}")
        except Exception as e:
            self.logger.error(f"发送微信通知异常: {e}", exc_info=True)
    
    def main(self):
        """主循环（轮询方法）"""
        self.logger.info("进入服务主循环（轮询方法）")
        
        try:
            # 验证配置是否有效
            check_interval = self.config.get("check_interval", 2)
            if not isinstance(check_interval, int) or check_interval < 1 or check_interval > 60:
                self.logger.error(f"无效的检测间隔配置: {check_interval}，使用默认值2秒")
                check_interval = 2
            
            while self.is_running:
                # 检查服务是否被请求停止
                rc = win32event.WaitForSingleObject(self.hWaitStop, 1000)
                if rc == win32event.WAIT_OBJECT_0:
                    self.logger.info("主循环收到停止信号，退出")
                    break
                
                try:
                    # 重新加载配置
                    self.config.load_config()
                    check_interval = self.config.get("check_interval", 2)
                    
                    # 检查锁屏状态
                    locked = self.is_screen_locked()
                    self.logger.debug(f"当前锁屏状态: {locked}")
                    
                    # 处理状态变化
                    self.lock_state_changed(locked)
                    
                    # 休眠指定间隔
                    self.logger.debug(f"主循环休眠 {check_interval} 秒")
                    time.sleep(check_interval)
                except Exception as e:
                    self.logger.error(f"主循环迭代异常: {e}", exc_info=True)
                    # 出错后等待更长时间，避免频繁重试
                    time.sleep(5)
        except Exception as e:
            self.logger.critical(f"主循环致命错误: {e}", exc_info=True)
            # 致命错误时退出服务
            self.is_running = False
    
    def run_with_wmi(self):
        """使用WMI方法运行服务"""
        self.logger.info("进入服务主循环（WMI方法）")
        
        try:
            # 验证配置是否有效
            check_interval = self.config.get("check_interval", 2)
            if not isinstance(check_interval, int) or check_interval < 1 or check_interval > 60:
                self.logger.error(f"无效的检测间隔配置: {check_interval}，使用默认值2秒")
                check_interval = 2
            
            while self.is_running:
                # 检查服务是否被请求停止
                rc = win32event.WaitForSingleObject(self.hWaitStop, 1000)
                if rc == win32event.WAIT_OBJECT_0:
                    self.logger.info("主循环收到停止信号，退出")
                    break
                
                try:
                    # 重新加载配置
                    self.config.load_config()
                    check_interval = self.config.get("check_interval", 2)
                    
                    # 检查锁屏状态和获取当前用户
                    locked, user = self.is_screen_locked_wmi()
                    self.logger.debug(f"当前锁屏状态(WMI): {locked}, 用户: {user}")
                    
                    # 处理状态变化
                    self.lock_state_changed(locked, user)
                    
                    # 休眠指定间隔
                    self.logger.debug(f"主循环休眠 {check_interval} 秒")
                    time.sleep(check_interval)
                except Exception as e:
                    self.logger.error(f"主循环迭代异常: {e}", exc_info=True)
                    # 出错后等待更长时间，避免频繁重试
                    time.sleep(5)
        except Exception as e:
            self.logger.critical(f"WMI主循环致命错误: {e}", exc_info=True)
            # 致命错误时退出服务
            self.is_running = False


def run_service():
    """运行服务主函数"""
    if len(sys.argv) == 1:
        # 作为服务运行
        try:
            pythoncom.CoInitialize()  # 初始化COM库
            servicemanager.Initialize()
            servicemanager.PrepareToHostSingle(LockScreenService)
            servicemanager.StartServiceCtrlDispatcher()
        except Exception as e:
            # 捕获服务启动时的异常
            print(f"服务启动失败: {str(e)}")
            # 尝试记录到日志文件
            try:
                log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'service_startup_error.log')
                with open(log_file, 'a', encoding='utf-8') as f:
                    f.write(f"{time.ctime()} - 服务启动失败: {str(e)}\n")
            except:
                pass
    else:
        # 处理服务命令行参数
        win32serviceutil.HandleCommandLine(LockScreenService)

if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] in ['install', 'remove', 'start', 'stop', 'restart', 'status']:
        # 服务管理命令
        run_service()
    else:
        # 运行配置界面
        print("请使用命令行参数管理服务：")
        print("  install - 安装服务")
        print("  remove - 卸载服务")
        print("  start - 启动服务")
        print("  stop - 停止服务")
        print("  restart - 重启服务")
        print("  status - 查看服务状态")    