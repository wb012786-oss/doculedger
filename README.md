# 银行对账单一键标准化工具 (Bank Reconciler)

> 上传银行 PDF / Excel / CSV，自动转标准格式 Excel，带异常检测、双文件对账、OCR 扫描件识别。  
> **隐私优先**：处理完立即删除，不落库。可选推送汇总到企业微信 / 飞书。

## 功能（v1.1）

- 支持 **文本型 PDF**、**扫描件 PDF（OCR 中英）**、Excel、CSV
- 自动识别主流银行关键词
- 标准字段输出 + 异常提示（大额 / 重复 / 日期缺失）
- 上传两份文件自动对账差异
- **可选**：处理完成后推送摘要到企业微信 / 飞书群机器人（仅汇总数字，不含明细）
- 无账号、处理后立即删临时文件

## 快速部署到 InstaBox

```bash
npx instabox@latest sh

unzip 银行对账.zip -d bank-reconciler && cd bank-reconciler

sudo apt-get update -qq
sudo apt-get install -y -qq tesseract-ocr tesseract-ocr-chi-sim poppler-utils

python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt

UPLOAD_DIR=/tmp/bank_uploads uvicorn app.main:app --host 0.0.0.0 --port 8000
```

发布 8000 端口为 HTTPS 公网链接即可访问。

### Docker

```bash
docker build -t bank-reconciler .
docker run -d --name bank-recon -p 8000:8000 -e UPLOAD_DIR=/tmp/bank_uploads bank-reconciler
```

## 企业微信 / 飞书推送

1. 群里添加自定义机器人，复制 Webhook  
2. 网页上传区下方粘贴（可留空）  
3. 处理完成后只推送汇总（笔数、收支、异常数），**不含流水明细**

## 隐私

- 文件只存在临时目录，处理完即删  
- Webhook 仅当次使用，不存储  
- 推送内容为汇总数字

## License

MIT
