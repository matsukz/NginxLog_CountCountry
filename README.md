# NginxLog_CountCountry
MySQLに保存したNginxの国別アクセスログを集計して割合を出力してくれるやつ

## Dashboardの式
```MySQL
SELECT time AS '日時',remote_addr AS '接続元アドレス',country AS '国',if (status = 200,'許可',if (status = 403,'拒否','転送エラー')) AS '動作' FROM log.nginx_access_log ORDER BY time DESC LIMIT 12;
```
