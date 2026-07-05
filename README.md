# Word accumulation tool
单词积累工具(气死雷颜器)，用于完成雷颜布置的20篇20词积累

## 使用说明

1. 运行`main.exe`

2. 键入单词(支持多个)

  示例格式：

  ```text
  never gonna give you up
  ```
  ```text
  never,gonna,let,you,down
  ```
  ```text
  never
  gonna
  tell
  a
  lie
  and
  desert
  you
  ```
  程序会自动忽略大写字母、特殊符号等。

3. 调整释义数量(可不做调整，反正是打印)、并行线程数(1~10，越多越快，建议5)

4. 点击生成

5. 如果要求设置User Agent，请按照如下步骤获取：
  1. 打开浏览器
  2. 搜索`User Agent`
  3. 随便打开一个网站
  4. 寻找类似`Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36 Edg/149.0.0.0`的字符串，复制并粘贴入框中
  5. 如果不会，可以键入上述User Agent，不一定能通用
  6. 如果还是不会，随便输入一串文本(大于20个字符，仅限英文字母、数字、特殊符号，不能有汉字)
  7. 再要是不会，你就别用这个程序了：）

  注：

  * 输入的User Agent会保存在根目录下的`_internal\UA.txt`中，如果程序可以找到它，则下次使用可不用输入；同样，在4中输入的User Agent也会写入文件中

  * 源代码User Agent文件存储在根目录下`UA.txt`

6. 打开桌面生成的`单词积累.docx`，根据需要进行调整，然后就可打印(你想手抄也行，但谁会想呢？)。

## 声明

本程序仅供完成雷颜布置的20篇20词积累而编写。请勿向雷颜透露、将本程序另作他用等，最终解释权归作者本人所有。

![starbucks_icon](favicon.ico)

© anghenuan

---

图标有彩蛋：）
