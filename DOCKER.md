# Docker デプロイメントガイド

Nodenberg API Server を Docker でデプロイする方法を説明します。

---

## 📦 必要なもの

- Docker (version 20.10以上)
- Docker Compose (version 2.0以上)

**インストール確認:**
```bash
docker --version
docker compose version
```

---

## 🚀 クイックスタート

### 1. 環境変数の設定

`.env` ファイルを作成します：

```bash
# .env.example をコピー
cp .env.example .env

# APIキーを生成して設定
echo "API_KEY=$(openssl rand -base64 32)" >> .env
```

### 2. Docker Compose で起動

```bash
# ビルドと起動
docker compose up -d

# ログを確認
docker compose logs -f
```

### 3. 動作確認

```bash
# ヘルスチェック
curl http://localhost:3100/health

# 結果:
# {
#   "status": "ok",
#   "timestamp": "2025-12-14T...",
#   "service": "Nodenberg API Server",
#   "version": "0.0.1"
# }
```

---

## 🛠️ Docker Compose コマンド

### 起動・停止

```bash
# バックグラウンドで起動
docker compose up -d

# フォアグラウンドで起動（ログ表示）
docker compose up

# 停止
docker compose down

# 停止＋ボリューム削除
docker compose down -v
```

### ビルド

```bash
# イメージを再ビルド
docker compose build

# キャッシュなしでビルド
docker compose build --no-cache

# ビルドして起動
docker compose up -d --build
```

### ログ確認

```bash
# 全ログ表示
docker compose logs

# リアルタイムでログ表示
docker compose logs -f

# 最新100行のログ
docker compose logs --tail=100
```

### コンテナ管理

```bash
# コンテナ一覧
docker compose ps

# コンテナに入る
docker compose exec nodenberg-api bash

# コンテナの状態確認
docker compose ps -a
```

---

## 🔧 手動での Docker ビルド・実行

Docker Compose を使わず、手動で実行する場合：

### イメージのビルド

```bash
docker build -t nodenberg-api:latest .
```

### コンテナの起動

```bash
docker run -d \
  --name nodenberg-api \
  -p 3100:3100 \
  -e API_KEY=your-secret-api-key-here \
  -e PORT=3100 \
  nodenberg-api:latest
```

### ログ確認

```bash
docker logs -f nodenberg-api
```

### コンテナの停止・削除

```bash
docker stop nodenberg-api
docker rm nodenberg-api
```

---

## ⚙️ 環境変数

Docker Compose で使用する環境変数：

| 変数名 | デフォルト値 | 説明 |
|--------|-------------|------|
| `PORT` | `3100` | APIサーバーのポート番号 |
| `API_KEY` | `default-secret-key-please-change-this` | API認証キー（必須） |
| `NODE_ENV` | `production` | 実行環境 |

### 環境変数の設定方法

**方法1: `.env` ファイル**
```bash
# .env ファイルを編集
PORT=3100
API_KEY=your-secret-api-key-here
NODE_ENV=production
```

**方法2: `docker-compose.yml` で直接指定**
```yaml
environment:
  - PORT=3100
  - API_KEY=your-secret-api-key-here
```

**方法3: コマンドラインで指定**
```bash
API_KEY=your-key docker compose up -d
```

---

## 🔒 セキュリティ設定

### APIキーの生成

強力なAPIキーを生成します：

```bash
# Base64エンコード（32バイト）
openssl rand -base64 32

# Hexエンコード（32バイト）
openssl rand -hex 32
```

### .env ファイルの保護

```bash
# .env ファイルのパーミッションを制限
chmod 600 .env

# Gitにコミットしないよう確認
cat .gitignore | grep .env
```

---

## 📊 ヘルスチェック

Docker Compose には自動ヘルスチェックが設定されています：

```yaml
healthcheck:
  test: ["CMD", "curl", "-f", "http://localhost:3100/health"]
  interval: 30s
  timeout: 10s
  retries: 3
  start_period: 40s
```

### ヘルスステータスの確認

```bash
# コンテナのヘルスステータスを確認
docker compose ps

# 詳細なヘルスチェック情報
docker inspect nodenberg-api | jq '.[0].State.Health'
```

---

## 🐛 トラブルシューティング

### コンテナが起動しない

```bash
# ログを確認
docker compose logs

# コンテナの詳細情報
docker inspect nodenberg-api
```

**よくある原因:**
- ポート3100が既に使用されている
- API_KEYが設定されていない
- Dockerデーモンが起動していない

### ポート競合の解決

```bash
# ポート3100を使用中のプロセスを確認
lsof -i :3100
sudo netstat -tulpn | grep 3100

# docker-compose.yml でポートを変更
ports:
  - "3200:3100"  # ホスト側を3200に変更
```

### LibreOfficeエラー

```bash
# コンテナに入ってLibreOfficeを確認
docker compose exec nodenberg-api bash

# LibreOfficeバージョン確認
soffice --version

# LibreOffice再初期化
node scripts/init-libreoffice.js
```

### メモリ不足

```bash
# docker-compose.yml にメモリ制限を追加
services:
  nodenberg-api:
    deploy:
      resources:
        limits:
          memory: 2G
        reservations:
          memory: 1G
```

---

## 📈 本番環境へのデプロイ

### 推奨設定

```yaml
# docker-compose.prod.yml
version: '3.8'

services:
  nodenberg-api:
    build: .
    container_name: nodenberg-api-prod
    ports:
      - "3100:3100"
    environment:
      - NODE_ENV=production
      - PORT=3100
      - API_KEY=${API_KEY}
    restart: always
    deploy:
      resources:
        limits:
          memory: 2G
          cpus: '1.0'
    healthcheck:
      test: ["CMD", "curl", "-f", "http://localhost:3100/health"]
      interval: 30s
      timeout: 10s
      retries: 3
      start_period: 40s
    logging:
      driver: "json-file"
      options:
        max-size: "10m"
        max-file: "3"
```

### 本番環境での起動

```bash
# 本番用のcomposeファイルを使用
docker compose -f docker-compose.prod.yml up -d

# または環境変数で指定
NODE_ENV=production docker compose up -d
```

---

## 🔄 更新とメンテナンス

### アプリケーションの更新

```bash
# 1. 最新コードを取得
git pull

# 2. イメージを再ビルド
docker compose build --no-cache

# 3. コンテナを再起動
docker compose down
docker compose up -d

# 4. ログで起動確認
docker compose logs -f
```

### バックアップ

```bash
# 環境設定のバックアップ
cp .env .env.backup

# Dockerイメージのエクスポート
docker save nodenberg-api:latest > nodenberg-api.tar

# Dockerイメージのインポート
docker load < nodenberg-api.tar
```

---

## 📡 外部からのアクセス設定

### リバースプロキシ (Nginx)

```nginx
# /etc/nginx/sites-available/nodenberg-api
server {
    listen 80;
    server_name api.yourdomain.com;

    location / {
        proxy_pass http://localhost:3100;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection 'upgrade';
        proxy_set_header Host $host;
        proxy_cache_bypass $http_upgrade;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

### SSL/TLS 設定 (Let's Encrypt)

```bash
# CertbotでSSL証明書取得
sudo certbot --nginx -d api.yourdomain.com
```

---

## 🧪 テスト

### コンテナ内でテスト実行

```bash
# コンテナに入る
docker compose exec nodenberg-api bash

# テスト実行
npm test
```

### 外部からAPIテスト

```bash
# APIキーを環境変数に設定
export API_KEY="your-secret-api-key"

# ヘルスチェック
curl http://localhost:3100/health

# テンプレート情報取得（要APIキー）
curl -X POST \
  -H "Content-Type: application/json" \
  -H "X-API-Key: $API_KEY" \
  -d '{"templateBase64":"..."}' \
  http://localhost:3100/template/info
```

---

## 📚 関連ドキュメント

- [API.md](API.md) - API仕様書
- [README.md](README.md) - プロジェクト概要
- [tests/README.md](tests/README.md) - テストガイド

---

## 💡 ベストプラクティス

### DO ✅

- APIキーは環境変数で管理する
- 本番環境では `restart: always` を設定する
- ログのローテーション設定を行う
- ヘルスチェックを有効にする
- メモリ制限を設定する

### DON'T ❌

- APIキーをハードコードしない
- `.env` ファイルをGitにコミットしない
- rootユーザーでコンテナを実行しない
- 本番環境で `latest` タグを使用しない
- ログを無制限に溜めない

---

## 🆘 サポート

問題が発生した場合:

1. [GitHub Issues](https://github.com/nodenberg/nodenberg/issues) で既知の問題を確認
2. `docker compose logs` でログを確認
3. `docker compose ps` でコンテナの状態を確認
4. 必要に応じて Issue を作成

---

最終更新日: 2025-12-14
