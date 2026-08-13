请帮 me在当前工程中自动搭建一套“实装窗口与 Review 窗口自动化双进程协作环境”，并将模型分别配置为 DeepSeek v4 Flash xmax 和 DeepSeek v4 Pro xmax。请在当前工程目录下自动创建并配置以下文件：

1. 创建配置目录 `automation_profiles/impl` 和 `automation_profiles/review`：
   - 创建 `automation_profiles/impl/opencode.json`：
     - `"model"` 设置为 `"deepseek-v4flash-xmax"`（实装阶段使用高速度模型）
     - 权限配置：`"edit": "allow", "bash": "allow", "read": "allow"`
     - System Prompt 设为：“你是高级开发工程师，负责根据 .opencode/task_bridge.json 中的 task_description 需求和 review_criteria 标准进行代码修改。修改完成后，将总结写入 task_bridge.json 中的 impl_summary 字段，并将 current_turn 修改为 'review'。”
   
   - 创建 `automation_profiles/review/opencode.json`：
     - `"model"` 设置为 `"deepseek-v4pro-xmax"`（Review 阶段使用高准确度模型）
     - 权限配置：`"edit": "deny", "bash": "deny", "read": "allow"`（严格禁止修改代码）
     - System Prompt 设为：“你是资深 Code Reviewer，禁止修改代码。负责检查 Git 改动，并严格对比 .opencode/task_bridge.json 中的 task_description、impl_summary 以及 review_criteria 中的每一条标准。若全部通过则将 current_turn 改为 'impl' 且 status 改为 'DONE'；若有任意不通过，请将具体退回意见写入 review_feedback 并将 current_turn 改为 'impl'。”

2. 在当前项目的 `.opencode/` 目录下创建通信交换文件 `task_bridge.json`：
   初始内容：
   {
     "current_turn": "impl",
     "status": "PENDING",
     "task_description": "请在这里填写你的初始开发需求",
     "review_criteria": [
       "1. 检查修改是否完整实现了 task_description 中的需求",
       "2. 检查是否存在未捕获的异常、死循环或内存泄漏风险",
       "3. 检查代码命名与风格是否规范，严禁硬编码敏感信息及魔法值"
     ],
     "impl_summary": "",
     "review_feedback": ""
   }

3. 在项目根目录创建双服务启动脚本 `start-services.bat`：
   写入 Windows 脚本，分别启动端口 3000 (指定 OPENCODE_CONFIG_DIR 为 automation_profiles/impl) 和端口 3001 (指定 OPENCODE_CONFIG_DIR 为 automation_profiles/review) 的 `opencode serve` 独立进程。

4. 在项目根目录创建看门狗脚本 `watchdog.ps1`：
   写入轮询逻辑，每隔 5 秒读取 `.opencode/task_bridge.json`。根据 current_turn 自动通过 REST API 向 3000 或 3001 端口发起任务（带有 180 秒 Timeout 超时保护和自动创建新 Session 机制）。

请开始自动生成以上所有文件，完成后通知我。
