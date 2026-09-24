# GitHub Pages 與中央試算 API 部署

GitHub Pages 只發布前端。建議書、商品版本與 LibreOffice 重算服務必須部署在私有的中央 API。

## 資料安全

- 不要把建議書 `.xlsx`／`.xlsm` 上傳至此公開 repository。
- 中央 API 將來源建議書存於私有 Cloud Storage bucket 掛載的 `/data`。
- `ADMIN_UPLOAD_TOKEN` 只設定在中央 API 的 Secret Manager，不可寫入 `app-config.js` 或 Git repository。

## Google Cloud Run 參考部署

準備一個 Google Cloud project、一個私有 Cloud Storage bucket，以及供 API 使用的 service account。該 service account 需要 bucket 的 `Storage Object User` 權限。

先建立 `ADMIN_UPLOAD_TOKEN` secret，然後將本 repository 建成 container image 並部署。以下變數需要替換為實際值：

```zsh
PROJECT_ID="your-gcp-project"
REGION="asia-east1"
BUCKET="your-private-insurance-proposals"
IMAGE="$REGION-docker.pkg.dev/$PROJECT_ID/insurance/insurance-api:latest"

gcloud builds submit --tag "$IMAGE"
gcloud run deploy insurance-financing-api \
  --image "$IMAGE" \
  --region "$REGION" \
  --allow-unauthenticated \
  --max-instances 1 \
  --memory 1Gi \
  --timeout 60 \
  --set-env-vars "INSURANCE_DATA_DIR=/data,CORS_ALLOWED_ORIGINS=https://anzuanzu.github.io" \
  --set-secrets "ADMIN_UPLOAD_TOKEN=insurance-admin-upload-token:latest" \
  --add-volume "name=product-data,type=cloud-storage,bucket=$BUCKET" \
  --add-volume-mount "volume=product-data,mount-path=/data"
```

`--max-instances 1` 在目前 JSON 商品清單的檔案式儲存模型下避免同時匯入造成衝突。若未來需大量同時匯入，應把商品清單與版本紀錄改存至交易式資料庫。

部署後取得 Cloud Run HTTPS URL，填入根目錄 `app-config.js`：

```js
window.INSURANCE_API_BASE_URL = "https://your-api-url";
```

提交並推送此設定後，GitHub Pages 使用者將讀取中央 API 的商品清單與即時計算結果。管理者上傳時會被要求輸入上傳金鑰；該金鑰只會暫存於瀏覽器 session，不會寫入 GitHub Pages。

## 首次商品匯入

1. 開啟 GitHub Pages 網站並確認中央 API 的 `/api/health` 可回應。
2. 使用具上傳金鑰的管理者帳號，逐一上傳五份檔名含「保費融資」的建議書。
3. API 會把來源檔與 `products.json` 存到私有 bucket。
4. 所有使用者重新整理頁面後，即會讀取相同的商品版本。

## 驗收

```zsh
curl https://your-api-url/api/health
curl https://your-api-url/api/products
```

確認網站可讀取商品後，以 ABA 預設條件試算，應顯示「第7年起持續增購保額」及第 20 年保障 216,607 USD。

