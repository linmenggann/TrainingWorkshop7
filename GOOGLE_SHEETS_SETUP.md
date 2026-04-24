# Google Sheets 串接設定

## 試算表資訊

- Google Sheets 網址：https://docs.google.com/spreadsheets/d/1vdhgjk6lzSUeP1O-CC73LPAFSpLJlT7E3nKx7At7JMA/edit?gid=0#gid=0
- 分頁名稱：`工作坊報名資料`

## 表頭

請在 `工作坊報名資料` 分頁第 1 列放入以下表頭。Apps Script 也會在分頁空白時自動建立相同表頭。

| A | B | C | D | E | F | G | H | I | J |
|---|---|---|---|---|---|---|---|---|---|
| 提交時間 | 姓名 | 機構名 | 職稱 | 負責職類 | 是否擔任教學訓練計畫主持人 | 預計參與方式 | Email | 聯繫電話 | 來源頁面 |

## Apps Script 部署步驟

1. 開啟 Google Sheets。
2. 點選「擴充功能」->「Apps Script」。
3. 將 `google-apps-script/Code.gs` 的內容貼入 Apps Script 編輯器。
4. 點選「部署」->「新增部署作業」。
5. 類型選擇「網頁應用程式」。
6. 執行身分選擇「我」。
7. 存取權限依需求選擇「任何人」或「知道連結的任何人」。
8. 部署後複製 Web App URL。
9. 回到 `index.html`，將 `GOOGLE_APPS_SCRIPT_URL` 的值改成 Web App URL。

## 前端送出欄位對照

| 前端欄位 name | Google Sheets 表頭 |
|---|---|
| `name` | 姓名 |
| `organization` | 機構名 |
| `jobTitle` | 職稱 |
| `profession` | 負責職類 |
| `isHost` | 是否擔任教學訓練計畫主持人 |
| `participationType` | 預計參與方式 |
| `email` | Email |
| `phone` | 聯繫電話 |
| `sourcePage` | 來源頁面 |
