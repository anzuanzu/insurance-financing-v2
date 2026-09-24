# GitHub Pages 與中央試算 API 部署

正式容器使用 Debian 映像檔提供的 `/usr/bin/soffice` 執行 LibreOffice Calc。

GitHub Pages 只發布前端。建議書、商品版本與 LibreOffice 重算服務必須部署在私有的中央 API。

## 資料安全

- 不要把建議書 `.xlsx`／`.xlsm` 上傳至此公開 repository。
- 中央 API 將來源建議書、商品版本與公開試算資料檔存於私有 Cloud Storage bucket 掛載的 `/data`。
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
  --max-instances 3 \
  --min-instances 0 \
  --concurrency 1 \
  --memory 1Gi \
  --timeout 60 \
  --set-env-vars "INSURANCE_DATA_DIR=/data,CORS_ALLOWED_ORIGINS=https://anzuanzu.github.io" \
  --set-secrets "ADMIN_UPLOAD_TOKEN=insurance-admin-upload-token:latest" \
  --add-volume "name=product-data,type=cloud-storage,bucket=$BUCKET" \
  --add-volume-mount "volume=product-data,mount-path=/data"
```

每個 Cloud Run 執行個體一次只處理一筆請求（`--concurrency 1`），避免 LibreOffice Calc 在同一容器內互相干擾；最多三個執行個體可平行處理三筆不同的新條件試算。`--min-instances 0` 不會產生待命執行個體費用，但首次請求可能有冷啟動時間。

每一筆完成的試算結果會以獨立資料檔發布，並在下一次讀取時合併為商品的共用快速資料；因此多個執行個體可以安全地同時發布不同條件的結果。管理者匯入建議書仍應一次只進行一筆，避免兩個管理者同時更新同一商品。

部署後取得 Cloud Run HTTPS URL，填入根目錄 `app-config.js`：

```js
window.INSURANCE_API_BASE_URL = "https://your-api-url";
```

提交並推送此設定後，GitHub Pages 使用者將讀取中央 API 的商品清單與即時計算結果。管理者上傳時會被要求輸入管理上傳密碼；該密碼只會暫存於瀏覽器 session，不會寫入 GitHub Pages。

## 首次商品匯入

1. 開啟 GitHub Pages 網站並確認中央 API 的 `/api/health` 可回應。
2. 使用具管理上傳密碼的管理者帳號，逐一上傳五份檔名含「保費融資」的建議書。
3. API 會以 LibreOffice Calc 驗證來源預設情境，將來源檔、`products.json` 與版本化商品資料 JS 存到私有 bucket。
4. 所有使用者重新整理頁面後，網頁會在背景載入新版商品資料，直接顯示新條件的數字；不會下載任何檔案。
5. 使用者第一次輸入尚未快取的特殊條件時，API 會重算一次並加入共用商品資料；之後相同條件的所有使用者會直接使用已驗證結果。

## 驗收

```zsh
curl https://your-api-url/api/health
curl https://your-api-url/api/products
```

確認網站可讀取商品後，以 ABA 預設條件試算，應顯示「第7年起持續增購保額」及第 20 年保障 216,607 USD。
