# 家計LINEボット（GAS版）

Google Apps Script（GAS）で動作する、家計管理LINEボットです。

Gmail通知・PayPay通知・LINEコマンドを連携し、
残高管理・月次集計・目標管理を自動化します。

(構成図)[./構成図,img]

---

## 📁 ファイル構成

```
/  
├── main.gs  
├── 構成図.png  
└── README.md 
```

---

## 🏗 システム構成

構成図は `構成図.png` を参照してください。

### 使用サービス

- Google Apps Script
- Gmail
- Google Spreadsheet
- LINE Messaging API
- MacroDroid（PayPay通知転送）

---

## ⚙ 仕組み概要

### ① Gmail監視（5分トリガー）

- 住信SBIデビット利用メールを取得
- PayPay出金メールを取得
- スプレッドシートへ保存
- 残高更新
- LINE通知

---

### ② PayPay通知（スマホ → Webhook）

MacroDroidからWebhook送信：

- PayPayチャージ通知を受信
- スプレッドシートへ保存（pending）
- 銀行出金メールと±1分で突合
- 一致したら残高更新＆LINE通知

---

### ③ LINEコマンド

| コマンド | 内容 |
|----------|------|
| メニュー | コマンド一覧表示 |
| 残高 | 現在の残高表示 |
| 更新 | 残高を上書き |
| 入金 | 残高に加算 |
| 今月 | 月次利用額集計 |
| 目標 | 目標管理モード |

---

## 📊 月次集計仕様

- `transactions` シートを集計
- 今月の利用額合計
- 店舗別ランキング（上位8件）
- PayPayチャージも含む

---

## 💾 スプレッドシート構成

### transactions
| ts | merchant | amount | currency | approval | message_id | source |

### paypay_events
| ts | amount | raw_text | status |

### bank_events
| ts | type | message_id | status | amount | raw_text |

### settings
| key | value |

---

## 🔐 Script Properties 必須設定

以下を設定してください：

- `LINE_CHANNEL_ACCESS_TOKEN`
- `GROUP_ID`

初回実行時に `SPREADSHEET_ID` は自動生成されます。

---

## 🔁 トリガー設定

- `pollGmail()` を5分間隔で実行

---

## 💰 コスト

- GAS：無料枠内
- Gmail：無料
- Spreadsheet：無料
- LINE：月200通まで無料

---

## 🚀 デプロイ手順

1. GASに `main.gs` を貼り付け
2. Webアプリとしてデプロイ
3. LINE Webhook URL設定
4. Script Properties設定
5. トリガー設定（5分）

---

## 🧠 実現できていること

- デビット自動記録
- PayPay出金自動突合
- 自動残高管理
- 月次レポート
- 目標管理
- LINE通知

---

## 🛠 メンテナンス

- 月次リセット：`resetThisMonthOnly()`
- 完全初期化：`resetAllFinanceData()`

---

## 注意

- LINE無料枠（200通/月）超過に注意
- Gmail仕様変更時は修正が必要
