# LINE webhook proxy

這是一個給你現在這份 `code.gs` 用的最小中介 webhook。

用途只有 3 件事：

1. 接住 LINE webhook，立刻回 `200 OK`
2. 解析你傳的日期，例如 `今天`、`明天`、`5/8`
3. 去呼叫你的 GAS `getLineDigest`，再用 LINE Push 回你

## 為什麼這樣比較穩

因為 LINE Verify 跟 webhook 比較喜歡打到「一般後端服務」，不喜歡碰到 Apps Script Web App 可能出現的 redirect 或部署版本問題。

## 你要先做的事

1. 把 `班級經營桌面工具/code.gs` 重新部署成 Web App
2. 存取權設成：
   - `Execute as`: 你自己
   - `Who has access`: Anyone
3. 記下 Web App 的 `/exec` 網址

## 正式部署到 Render

如果你不想一直開著本機的 `npm start` 和 `cloudflared`，可以直接部署到 Render。

這個資料夾已經附好：

- [render.yaml](./render.yaml)

### 部署步驟

1. 先把整個專案推到 GitHub
2. 到 Render 建立帳號並連結 GitHub
3. 在 Render 選 `New` -> `Blueprint`
4. 選你的 GitHub repo
5. Render 會讀到這份 `render.yaml`
6. 建立服務後，到 Render 的環境變數頁面填入：

```text
LINE_CHANNEL_SECRET
LINE_CHANNEL_ACCESS_TOKEN
LINE_DEFAULT_USER_ID
GAS_BASE_URL
```

### Render 部署後要做的事

1. 找到 Render 給你的網址，例如：

```text
https://line-webhook-proxy.onrender.com
```

2. 到 LINE Developers Console，把 `Webhook URL` 改成：

```text
https://line-webhook-proxy.onrender.com/webhook
```

3. 按 `Verify`

### 補充

- `LINE_DEFAULT_USER_ID` 可以先留空
- `GAS_BASE_URL` 要填你 GAS Web App 正式部署的 `/exec` 網址
- Render 免費方案有可能在閒置後休眠，第一次訊息可能稍慢

## 本機測試

在這個資料夾建立 `.env`，內容可以直接照 `.env.example` 改：

```env
PORT=8787
LINE_CHANNEL_SECRET=你的 channel secret
LINE_CHANNEL_ACCESS_TOKEN=你的 long-lived channel access token
LINE_DEFAULT_USER_ID=你的 userId
GAS_BASE_URL=https://script.google.com/macros/s/你的部署ID/exec
```

啟動：

```powershell
cd "c:\Users\asin6\.cursor-tutor\班級經營桌面工具\line-webhook-proxy"
npm start
```

健康檢查：

```powershell
curl http://localhost:8787/health
```

## 給 LINE 的 webhook URL

如果只是本機測試，要再配一個公開網址工具，例如：

- `cloudflared tunnel --url http://localhost:8787`
- 或 `ngrok http 8787`

然後把公開網址後面的 `/webhook` 填進 LINE Developers Console。

範例：

```text
https://xxxxx.trycloudflare.com/webhook
```

## 你可以怎麼測

1. 先在 LINE Developers 按 `Verify`
2. Verify 過了之後，在你的 LINE 官方帳號聊天室傳：
   - `今天`
   - `明天`
   - `5/8`
3. 看它有沒有回你教學進度、聯絡簿、本週記事

## 備註

- 這個中介 webhook 會驗證 `x-line-signature`
- 它不直接碰 Google Sheet，只呼叫你 GAS 的 `getLineDigest`
- 如果你之後要正式上線，可以把這一包放到 Cloud Run、Railway、Render 之類的平台
