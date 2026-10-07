# 教务系统表单填写参考

本文件是第 3~5 步的浏览器操作细节。页面交互一律遵循 **browser-use-automation** 技能（`computer_use_tool`, plane="bu"，`import seed_browser_use as bu`，observe-act-observe 循环）。

## 入口与登录

- 教务系统网址以用户提供为准；默认入口：`http://jiaowunew.stiei.edu.cn/jwapp/sys/wdjxrwapp/*default/index.do`（进入"我的教学任务"）。
- 定位路径：课程列表 → 按班级/课程名找到目标课程 → 进入"教案"模块。
- **登录/验证码接管（强制）**：当页面出现登录、验证码、二次验证，或 `bu` 观察返回 `blocked=auth` 时，下一步必须调用 `interaction_request_action`：

```json
{
  "type": "browserControl",
  "display_message": "教务系统需要登录/验证码。请接管浏览器完成登录（可输入账号密码、拖动滑块等），完成后把控制权交回，我会重新读取页面继续。"
}
```

- 接管前确认登录页确实处于前台标签页；用户完成后，先 `bu.snapshot()` 或 `bu.page_info()` 重新观察、确认认证已通过，再继续。**不沿用接管前的任何 ref**。
- 不读取、不保存、不代输密码与验证码。

## 逐课次填写闭环（第 4 步核心）

每课次严格按以下顺序推进，处理完一课次再处理下一课次：

1. 在课程列表中按日期/周次定位该课次，点击"教案"链接
2. `bu.wait_for_load()` + `bu.snapshot()`，确认表单已加载（出现输入框/textarea）
3. **探测表单 DOM**（见下节），确认目标控件真实索引
4. 按字段映射赋值（见"字段映射与赋值"）
5. 勾选教学载体（第一组）/教学媒介（第二组）checkbox
6. 上传教学副页（从素材库选内容相近文件；PDF 优先）
7. 点击"暂存"（默认）或"提交"（用户明确要求时），等待"操作成功"提示
8. 关闭面板返回列表；若列表发生翻页，操作下一课次前先确认当前页与目标一致

## 表单 DOM 探测（每次打开表单必做，不写死索引）

- 用 `bu.read_all()` 或快照统计当前表单中可见的 `input`、`textarea`、`checkbox`，按出现顺序记录索引与可辨识上下文（placeholder、label、名称）。
- 常见形态：标题等短文本为 `input`，目标/重点/难点/实施等长文本为 `textarea`。默认经验映射（**仅作参考，以实测为准**）：目标输入框索引 13/14/15、textarea 0-4 依次对应教案内容字段；实测不一致时必须按实际 DOM 调整。
- 若表单控件无稳定 label/placeholder，用 `bu.screenshot(ref=...)` 查看控件附近文字辅助判断归属。
- 每次打开新表单都重新探测：不同课程/不同页签的表单结构可能不同。

## 字段映射与赋值

- 将教案表格中的字段写入对应控件：教案标题 → 标题输入框；知识目标/能力目标/素养目标/教学重点/教学难点/布置作业/教学实施/课程思政 → 对应 textarea。
- **事件派发（强制）**：框架类表单（vue/antd 等）仅设置 `value` 不生效，必须同步派发事件。用 `bu.js()` 窄范围执行：

```python
bu.js("""
  const el = document.querySelector('textarea:nth-of-type(2)'); // 示例：以实测选择器为准
  const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set;
  setter.call(el, '教案内容...');
  el.dispatchEvent(new Event('input', {bubbles: true}));
  el.dispatchEvent(new Event('change', {bubbles: true}));
""")
```

  或优先使用 `bu.fill_input(selector, text)` 等受支持接口；普通原生输入框可直接 `bu.type(ref, text)`。
- 赋值后回读一次（`bu.get(ref, "value")` 或快照）确认写入成功，再继续下一步。

## 教学载体 / 媒介勾选

- 第一组（教学载体）：默认勾选"教材"；实训课等按课程类型改为"教材 + 手册/工作页"。
- 第二组（教学媒介）：默认勾选"多媒体设备、实物"，避免勾选过多无关项。
- 用 `bu.snapshot()` 定位 checkbox ref，`bu.click(ref)` 勾选；勾选后回读 checked 状态确认。

## 教学副页上传

- 从用户指定的教学素材库目录中选择与当前课次主题内容相近的文件；**PDF 优先**，无对应 PDF 时可上传任意相近格式文件。
- 上传方式：`bu.upload(ref, absolute_path)` 或 `bu.upload_file(selector, absolute_path)`。
- **副页必填**：教学副页缺失会导致暂存失败；无精确匹配时用相近素材代替，不可跳过。
- 若上传需要用户在本机选择文件（无法定位本地路径），调用 `interaction_request_action`（type="fileUpload"）请用户选择，完成后重新观察继续。

## 暂存 / 提交

- **默认只暂存**：点击"暂存"按钮 → 等待"操作成功"提示（必要时 `bu.find("操作成功")` 或等待 toast 出现）→ 关闭面板。
- **提交（仅用户明确要求）**：回到列表 → 勾选所有未提交教案 → 点击"提交教案" → 处理确认弹窗（`bu.click_and_handle_dialog(ref, accept=True)` 或 `bu.handle_dialog`）→ 等待"操作成功"。
- 若点击后无反应：先快照检查是否有弹窗、校验错误或 toast，再对症处理；不盲目重复点击。

## 翻页与定位注意

- 列表翻页后可能回跳第 1 页：翻页后先确认当前页显示的目标班级/课次，再执行操作。
- 目标行定位用 `bu.read_all()` 按"日期/周次/主题"文本匹配，取文档级 ref；不凭记忆坐标点击。

## 常见失败恢复

| 现象 | 处理 |
|---|---|
| `BU_REF_STALE` / ref 失效 | 重新 `bu.snapshot()` 取新 ref |
| `BU_SESSION_STALE` | `bu.resync()` 后重新观察 |
| 表单赋值后保存丢内容 | 确认赋值时已派发 input/change 事件，且保存前回读 |
| 暂存失败提示缺项 | 检查必填字段与教学副页是否齐全，补齐后重试 |
| 登录/验证码/`blocked=auth` | `interaction_request_action`(browserControl) 接管，完成后重观察 |
| 点击无响应 | 快照检查弹窗/校验/toast，对症处理 |
