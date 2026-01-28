# Repository Guidelines

## 说明

-项目处于实验阶段，仅开源说明，后续根据需求更新

## 项目结构与模块组织

- `backend/` 为 FastAPI 服务端。
  - `backend/app/main.py` 启动应用与中间件。
  - `backend/app/routes/` 定义 API 路由。
  - `backend/app/services/` 负责 Excel 处理、会话存储、LLM 调用。
  - `backend/app/models/` 定义请求/响应模型。
- `frontend/` 为 Vite + React 前端。
  - `frontend/src/` 放 UI 组件与样式。
  - `frontend/public/` 存放静态资源。

## 构建、测试与开发命令

后端（PowerShell）：
```powershell
cd backend
pip install -r requirements.txt
uvicorn app.main:app --reload --port 8000
```

前端：
```powershell
cd frontend
npm install
npm run dev
```

常用命令：
- `npm run lint`：运行 ESLint。
- `npm run build`：构建前端生产包。

## 代码风格与命名规范

- Python：4 空格缩进，函数保持短小、可读，服务逻辑放在 `services/`。
- TypeScript/React：2 空格缩进，变量/函数用 `camelCase`，组件用 `PascalCase`。
- API 路由保持 REST 风格，如 `/sessions/{id}/workbooks/upload`。
- UI 修改前后建议运行 `npm run lint`。

## 测试指南

当前暂无测试套件。如果新增测试，建议：
- 后端：使用 `pytest`，放在 `backend/tests/`，文件名 `test_*.py`。
- 前端：使用 `vitest` 或 `@testing-library/react`，放在 `frontend/src/__tests__/`。

新增测试命令请更新本文件。

## 提交与 PR 规范

项目已建立 git 历史，暂无既有提交规范。如初始化 git，建议使用清晰的动词开头，例如 “Fix preview refresh on upload”。

如后续启用 PR：
- 提供简短说明、关键 UI 截图、可复现步骤。
- 关联需求或问题编号。

## 安全与配置

- API Key 放在 `backend/.env`（参考 `backend/.env.example`）。
- 不要提交密钥，`.env` 仅用于本地。
- 部署前检查 CORS 和 API 地址配置。
