# 沈阳基地服务沟通看板自动化部署指南

## 📦 仓库新增内容

```
service-kanban/
├── scripts/
│   ├── download.py      # 用 cookie 下载当日 Excel
│   └── preprocess.py    # Excel → kanban_data.json
├── .github/workflows/
│   └── update-dashboard.yml  # 每天 9:00 AM 自动跑
└── tools/
    └── export-cookies.html   # 一键导出 cookie 的本地工具
```

## 🚀 一次性配置(约 10 分钟)

### 第 1 步:把脚本推到 GitHub

方法 1:用我给你的 PAT(最简单,我直接推)
方法 2:手动上传以下 4 个文件到 GitHub:
- `scripts/download.py`
- `scripts/preprocess.py`
- `.github/workflows/update-dashboard.yml`
- `tools/export-cookies.html`

### 第 2 步:导出 Cookie

1. 在 Chrome 打开 [https://xwv5.aidoutang.com/dwpush/](https://xwv5.aidoutang.com/dwpush/)
2. 登录 SSO(走完图形验证码)
3. 把标签页 URL 改成 `https://dw.aidoutang.com/` 并回车
4. 按 `F12` → `Console` 标签
5. 粘贴运行:
   ```javascript
   navigator.clipboard.writeText(document.cookie).then(() => alert('已复制 ' + document.cookie.length + ' 字符'));
   ```
6. 弹窗说"已复制"就行

### 第 3 步:在 GitHub 配置两个 Secret

打开 [https://github.com/mjx0217/service-kanban/settings/secrets/actions](https://github.com/mjx0217/service-kanban/settings/secrets/actions)

| Name | Value |
|------|-------|
| `AIDOUTANG_COOKIE` | 第 2 步剪贴板里的字符串 |
| `PUSH_TASK_ID` | `3084` |

### 第 4 步:手动触发一次测试

打开 [https://github.com/mjx0217/service-kanban/actions/workflows/update-dashboard.yml](https://github.com/mjx0217/service-kanban/actions/workflows/update-dashboard.yml)

点 `Run workflow` → 选 main → 绿色按钮

等 2-3 分钟看是否成功。

## 🔄 自动化运行

- **每天 9:00 AM 北京时间** 自动触发
- 如果 Cookie 失效,workflow 会失败并提示
- 重新导出 Cookie → 更新 Secret → 手动再触发一次

## ⚠️ 已知限制

- `replyRate` 和 `avgRespTime` 字段始终为 0(底表 Excel 不含这两个字段,如需启用需另找数据源)
- 0.1% 边界舍入差异(从 10818 个字段里约 18 个,与 OpenClaw 原数据有 0.1% 偏差,可忽略)
- 如果 `id=3084` 这个任务 ID 每天变,需要每天动态获取(我后面再优化)
