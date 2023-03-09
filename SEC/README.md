<h1>「SEC CIK Automatic CrawlSearch & CompName Similarity Comparison」</h1>

CIK為Central Index Key，為SEC電腦系統中識別已向 SEC 提交披露的公司和個人代碼


而此程式目的為

(1)年報所抓取非統一格式公司名稱(ex:7-eleven Inc)也能成功從SEC查詢系統搜索該企業CIK(可能資料庫為7-ELEVEN inc)

(2)年報所抓取非統一格式公司名稱成功比對Compustat Capital IQ而獲取CIK

因而處理格式不一問題，並進行字串比對

過程中(2)演算法能套用至(1)，(1)目前為刪除暫留字並進行搜索，而(2)套用至(1)演算法或許能最佳化搜索結果但尚未測試，
而(2)能調整正確比率，目前為正確率50%為最低標準，結果以降冪排序因此能找到最佳結果，兩者比對字串也不須長度一樣。
![image](https://user-images.githubusercontent.com/52849538/224116330-c1644e98-9eb0-4459-8b79-3468c16d2f49.png)
