<h1>1. 原始数据及输出数据</h1>
<h2>1) 原始数据</h2>
原始数据为一个.csv文件，其中包含了学者的姓名、单位和研究方向，要求根据此信息补全学者的其他信息。

<h2>2）输出文件</h2>
本程序的输出文件为一个工作簿（即.xlsx文件）。
该工作簿的内容包括了：学者的姓名、单位、研究方向、email、总被引次数、hindex、hindex5y、i10index、i10index5y、照片url、电话、地址、职称、国家。
其中学者的姓名、单位、研究方向是输入文件给出，email、总被引次数、hindex、hindex5y、i10index、i10index5y、照片url从谷歌学术页面中爬取得来，电话、地址、职称、国家则是从谷歌搜索结果摘要中爬取。

<h1>2. 程序说明</h1>
本程序从谷歌学术和谷歌搜索中获得学者的信息，具体涉及的网址较多，请在scholar.py和Google_complement.py中查看。
本程序在代码中已有各语句详细说明故不在此赘述，以下只给出主要部分代码的功能。

scholarly.py: python包scholar.py修改版本，省去了不必要页面的爬取，并支持ip代理。

inuput.py:将提取的结果写入到result.xlsx文件。

record.py:将已爬取的信息记录下来避免重复爬取。

complement.py: 学者信息补全主程序，使用了多进程提高爬取效率
运行示例: python complement(_noip).py [开始id] [结束id] [进程数]


