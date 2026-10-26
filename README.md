# Sheet Guardian

一个基于 Vue 3、TypeScript、Vite 和 Ant Design Vue 的 Chrome Manifest V3 扩展。它使用**当前 Chrome 配置文件中的 Google 主账号**访问 Google Drive，批量治理 Drive 文件共享权限，并为 Google 表格提供额外的 Sheet 保护权限同步。

## 已实现功能

- 点击扩展图标，在新标签页打开完整管理界面。
- 扫描当前 Google 账号可访问的全部 Drive 文件，而不只限于表格；支持 Google 表格、文档、幻灯片、文件夹、PDF、图片与其他文件。
- 自动识别当前 Google 账号拥有的文件，可按文件类型、所有权、公开链接、域内共享或受限状态筛选，并跨分页选择全部自有文件。
- 扫描结果可按文件名、文件 ID、所有者姓名/邮箱搜索，并支持最新/最旧修改日期排序。
- 扫描列表每页可显示 10、20、50、100、500 或 1000 条记录。
- 批量勾选后可一键取消全部选择，包括其他分页中保留的勾选项。
- 扫描结果可批量删除文件：采用移入 Google Drive 回收站的可恢复方式，执行前必须二次确认，无权操作的文件会跳过并记录失败。
- 批量关闭“知道链接的任何人”及 Workspace 域共享。
- 扫描结果中可批量开启“知道链接的任何人”，统一设置可查看、可评论或可编辑，并在执行前进行风险确认。
- 收紧共享时保留所有明确添加的用户/群组账号和所有者权限。
- 粘贴多条表格链接或文件 ID，批量只读查询所有者、共享范围、当前账号管理能力及每个用户/群组的权限角色，并可复制汇总结果。
- 批量权限查询结果可按邮箱或 Google 群组自动归集表格链接；支持搜索邮箱并单独复制该邮箱的全部链接，也可一键复制当前登录邮箱读取成功的全部链接。
- 粘贴多条 Drive 文件链接或文件 ID，再粘贴多个邮箱，可选择“组合授权”执行“文件 × 邮箱”，或选择“一一对应”严格执行第 1 个文件 → 第 1 个邮箱、第 2 个文件 → 第 2 个邮箱。
- 一一对应模式保留输入顺序和重复项，并要求文件数与邮箱数完全相同，避免把全部账号交叉添加到全部文件。
- 批量授权页提供独立的“公开链接”模式：完全不需要填写邮箱，可直接把每个目标文件设置为“知道链接的任何人”可查看、可评论或可编辑；公开编辑默认显示高风险警告。
- 也可直接从扫描结果选中文件并进入批量授权，无需复制链接。
- 批量授权时可选择“同时添加到 Sheet 保护权限”：按自定义关键词匹配内部工作表标题，只向已有且可修改的保护规则追加同一批邮箱；未设置保护、仅警告保护或无权修改的规则自动跳过。
- 支持可查看、可评论、可编辑三种角色，以及 Google 通知邮件和自定义附言。
- 批量添加邮箱时可选择同时开启链接共享，并为公开链接单独设置查看、评论或编辑角色；默认关闭。
- 所有权转移同时支持组合模式（多个文件转移给同一账号）和一一对应模式（每个文件分别转移给同序号账号），并在执行前展示配对预览与二次确认。
- 自动解析、校验、去重；已有相同权限跳过，已有不同权限更新角色。
- 执行前显示预估操作数，执行中显示进度，失败不阻塞其他任务。
- 最多保留 500 条本机操作日志，并可导出 CSV。
- 网页开发环境自动进入示例数据模式，不会修改真实 Drive 数据。

## 本地开发

```bash
npm install --cache ./work/npm-cache
npm run dev
npm run test
npm run build
```

生产扩展输出在 `dist/`。

## 配置 Google OAuth

Chrome 扩展 OAuth 客户端不能安全地做成一个供所有安装者共用的占位凭据，因此第一次使用需要配置自己的 Google Cloud 项目。

1. 先运行 `npm run build`。
2. 打开 `chrome://extensions`，开启“开发者模式”，选择“加载已解压的扩展程序”，加载本项目的 `dist/`。
3. 在扩展详情中复制“扩展程序 ID”。同一路径下重复构建、重新加载时，该 ID 会保持不变。
4. 在 Google Cloud Console 新建或选择项目，同时启用 **Google Drive API** 和 **Google Sheets API**。
5. 配置 OAuth 同意屏幕：
   - 个人 Gmail 可选 External，并把实际使用账号加入 Test users；
   - Google Workspace 内部工具可按组织策略选择 Internal；
   - 添加 `https://www.googleapis.com/auth/drive` scope。
6. 创建 OAuth Client ID，应用类型选择 **Chrome Extension**，填入第 3 步的扩展程序 ID。
7. 把得到的 Client ID 替换到 `public/manifest.json` 的 `oauth2.client_id`：

```json
"client_id": "你的客户端ID.apps.googleusercontent.com"
```

8. 再次运行 `npm run build`，回到 `chrome://extensions` 点击该扩展的“重新加载”。
9. 点击扩展图标，在新标签页中授权 Google 账号。

OAuth Client ID 不是客户端密钥；本扩展不需要、也不应保存 Client Secret。

## 权限说明

- `identity` / `identity.email`：通过 Chrome Identity 使用当前浏览器配置文件中的 Google 主账号。
- `https://www.googleapis.com/*`：调用 Google Drive API。
- `https://sheets.googleapis.com/*`：读取并更新匹配工作表中的已有保护规则。
- `https://www.googleapis.com/auth/drive`：扫描用户可访问文件，并创建、更新、删除共享权限。

`drive` 是高权限 OAuth scope。个人测试模式可直接使用；如果要面向大量外部用户发布，Google 可能要求完成 OAuth 应用验证。若只是组织内部使用，建议将 OAuth 应用限制为 Workspace Internal。

## 安全行为

- “收紧共享”只处理 `anyone` 和用户明确选中的 `domain` 权限，绝不删除 `user`、`group` 或 `owner` 权限。
- “同时添加到 Sheet 保护权限”默认关闭；启用后只更新已有保护规则的 `editors.users`，不会新建保护规则，也不会覆盖原有用户、群组或域设置。
- 所有权转移的组合模式只接受一个目标邮箱；一一对应模式允许多个目标邮箱，但必须与文件逐项等量配对。Workspace 同组织转移通常立即生效，原所有者会降为编辑者。
- “删除文件”只会将文件移入 Google Drive 回收站，不会清空回收站或永久删除；恢复与最终保留期限由 Google Drive 管理。
- 个人 Gmail 之间会先发送待接受的所有权转移邀请；目标账号接受前，原账号仍是所有者。
- 共享云端硬盘中的文件归组织所有，不支持转移给个人账号；跨 Workspace 组织转移也可能被组织策略拒绝。
- 每次收紧前都会重新读取权限明细；继承自上级文件夹或共享云端硬盘的权限会自动跳过，并在日志中提示来源处理方式。
- Google Drive 的 limited access 只适用于文件夹，不能直接应用到 Google 表格文件；表格的继承权限需要在其上级文件夹或共享云端硬盘调整。
- “收紧共享”提供默认关闭的高风险选项：遇到继承权限时，可自动为表格的直属父文件夹启用 limited access。该设置会阻断文件夹中的全部继承访问（包括继承的用户和群组），并会先检查当前账号是否具备 `canDisableInheritedPermissions` 能力。
- 没有 `canShare` 能力的文件不会执行共享修改；如果仍具备 `canTrash`，可单独用于移入回收站。批量链接中的无权限文件会记录失败而非反复请求。
- 所有 API 调用直接从扩展发往 Google，不经过自建服务器。
- 操作记录只保存在当前扩展的本地空间，可随时清除或导出。

## 项目结构

```text
public/manifest.json       Chrome MV3 清单与 OAuth scope
public/background.js       点击扩展图标时打开新管理页
public/icons/              Chrome 工具栏及扩展管理页图标
assets/brand/              高分辨率透明背景图标主稿
src/App.vue                管理台界面与批量工作流
src/lib/googleDrive.ts     Chrome Identity 与 Drive API 封装
src/lib/parsers.ts         链接、文件 ID、邮箱解析及并发控制
src/styles.css             响应式界面样式
```
