# qth_crawler
教务网爬虫，分为模拟登陆和验证码识别两个部分。<br />

爬虫使用的是urllib和requests库，难点其实在验证码识别用了80%的时间。

验证码识别调用Google的OCR模块，简单的验证码识别率70%左右。<br />

效率不高，只有每秒一个请求，请求次数过多服务器还会返回#10060 error，可能是服务器端的反爬虫策略。
