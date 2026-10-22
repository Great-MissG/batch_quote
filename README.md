# Fmile 批量询价

独立 Streamlit 应用，仅询价，不创建订单。

## 本机运行

在此目录执行 `启动.ps1`，打开 http://127.0.0.1:8501。
首次安装：Python 3.12+ 创建 `.venv`，执行 `python -m pip install -r requirements.txt`。
服务端 Token 读取顺序为环境变量、Streamlit Secrets、本目录 `.env`，字段名为 `FIMILE_API_TOKEN`。
不要提交 `.env`、`.streamlit/secrets.toml` 或 `.venv`。

## 业务规则

- A–G 为输入，上传文件只包含这7列即可；也接受完整模板。导出时自动补齐H–W结果列及表头格式。
- 每行一票；同一行的所有托盘尺寸和单托盘重量相同。不同规格请分票；当前模板不支持多规格合票。
- `units` 为托盘数；`totalWeight` 为托盘数乘单托盘重量。尺寸 inch、重量 lbs，整票必须大于20 lbs。
- 包装类型明确传 `Pallet`；`isStackable`、`isHazardous`、`isUsed`、`isMachinery` 均显式传布尔值 `false`。
- `pieces` 固定传整数 `1`；`units` 仍按上传的托盘数量传入。
- 5位邮编保持不变，ZIP+4仅取前5位；Excel数值邮编不足5位时补前导零。字符串非法邮编拒绝提交。
- 页面统一选择提货日期；货运等级通过建议等级接口获取。
- `accessorialServices: []`，接口仍固定加入 APPT。
- 每个 carrierCode 取最低报价，然后按接口原价升序选择4家。不减价，不强制包含 CEVA。
- 保险是否承保缺少明确依据时显示“未明确”，不把0保险金额或承运商责任限额当成承保结论。
- 时效保留接口天数和预计送达日期；null显示“未提供”，0天保留为0天。
- 结果保留原始输入和模板格式，最多4组报价；不足4家留空。失败行首组报价标注失败和原因。
- 页面只显示上传检查数量、询价进度、成功/失败数量和下载按钮，不展示输入或报价明细；错误原因可按需展开。Excel保留完整数据与全部报价。
- 一票失败不会终止其他票；401会停止后续提交。没有自动重试，已成功行不会因页面刷新交互重复提交。
- 上传文件、接口返回、Token不写入应用日志；上传与结果仅保留当前会话内存。关闭会话前下载结果。

## 上线

本机地址只在此电脑可用，不能作为公网链接。
Streamlit Community Cloud 需有 GitHub 代码仓库和你自己的 Streamlit 登录账号。
选择 `app.py` 为入口，Python 3.12；在平台 Secrets 设置 `FIMILE_API_TOKEN`。
应用访问权限设为私有，只授权自己的账号。不要通过公开网址替代访问控制。
本机启动脚本通过命令行绑定127.0.0.1；共享配置不固定监听地址，供Streamlit平台部署使用。部署其他服务器时应配置平台认证。
不包含凭据或上传文件的代码可部署；不要上传整个工作目录或虚拟环境。

接口基址：`https://oms.fimile.com/oms-manager/api/ltl`。
文档：https://www.fimile.com/ltlapidoc.html
