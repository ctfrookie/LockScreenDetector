锁屏检测服务 - Windows 系统锁屏状态监控与通知



功能特点





*   实时监控 Windows 系统锁屏 / 解锁状态


*   检测到状态变化时发送企业微信通知


*   支持获取当前登录用户信息并包含在通知中


*   作为 Windows 服务后台运行，支持开机自启动


*   提供灵活的配置选项和详细的日志记录


环境要求





*   Windows 7/8/10/11 系统


*   Python 3.6+


*   以下 Python 库:




```
pywin32


requests


wmi
```

安装与配置



### 1. 安装依赖库&#xA;



```
pip install pywin32 requests wmi
```

### 2. 下载并配置项目&#xA;



1.  下载`lock_screen_service.py`文件到本地


2.  创建`config.json`配置文件（首次运行会自动生成）


3.  编辑`config.json`设置企业微信 Webhook URL：




```
{


&#x20;   "webhook\_url": "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=你的Webhook密钥",


&#x20;   "check\_interval": 2,


&#x20;   "log\_level": "INFO",


&#x20;   "detection\_method": "wmi"


}
```

### 3. 安装为 Windows 服务&#xA;

以管理员身份打开命令提示符，执行以下命令：




```
\# 安装服务


python lock\_screen\_service.py install


\# 启动服务


python lock\_screen\_service.py start


\# 设置为自动启动


sc config LockScreenDetector start= auto
```

服务管理命令





```
\# 查看服务状态


python lock\_screen\_service.py status


\# 停止服务


python lock\_screen\_service.py stop


\# 重启服务


python lock\_screen\_service.py restart


\# 卸载服务


python lock\_screen\_service.py remove
```

配置说明



### config.json 配置项说明&#xA;



| 配置项&#xA;               | 类型&#xA;     | 默认值&#xA;    | 说明&#xA;                            |
| ---------------------- | ----------- | ----------- | ---------------------------------- |
| webhook\_url&#xA;      | string&#xA; | ""&#xA;     | 企业微信 Webhook 地址，用于发送通知&#xA;        |
| check\_interval&#xA;   | int&#xA;    | 2&#xA;      | 检测间隔时间（秒），范围 1-60&#xA;             |
| log\_level&#xA;        | string&#xA; | "INFO"&#xA; | 日志级别：DEBUG/INFO/WARNING/ERROR&#xA; |
| detection\_method&#xA; | string&#xA; | "wmi"&#xA;  | 检测方法：wmi（推荐）或 polling&#xA;         |

检测方法说明



### 1. WMI 方法（推荐）&#xA;

通过 Windows Management Instrumentation (WMI) 查询系统状态：




*   检查当前登录用户信息


*   检测锁屏界面进程 (LogonUI.exe) 存在与否


*   优点：不依赖窗口句柄，适合服务环境运行


### 2. 轮询方法&#xA;

通过窗口句柄检测锁屏状态：




*   获取当前前台窗口句柄和类名


*   检查是否为锁屏界面窗口类名


*   优点：检测更直接，缺点：可能受服务会话限制


日志查看



服务运行日志存储在脚本同目录下的`service.log`文件中，包含：




*   服务启动 / 停止记录


*   锁屏 / 解锁事件记录


*   错误异常信息


常见问题解决



### 1. 服务启动失败&#xA;



*   确保以管理员身份执行命令


*   检查日志文件`service_startup_error.log`获取详细错误


*   确认 Python 环境路径正确


### 2. 通知未发送&#xA;



*   检查`config.json`中 Webhook URL 是否正确


*   手动测试 Webhook：`curl -X POST -H "Content-Type: application/json" -d '{"msgtype":"text","text":{"content":"测试通知"}}' 你的Webhook地址`

*   查看日志中是否有发送通知的错误记录


### 3. 锁屏检测不准确&#xA;



*   尝试切换检测方法：修改`config.json`中的`detection_method`为`polling`

*   检查日志中锁屏检测的详细记录


*   确保服务运行在正确的用户会话中


卸载服务





```
python lock\_screen\_service.py stop


python lock\_screen\_service.py remove
```

安全提示





*   建议将脚本和配置文件存储在安全目录


*   企业微信 Webhook 密钥请妥善保管


*   服务运行需要管理员权限，请确保系统安全





> （注：文档部分内容可能由 AI 生成）
>