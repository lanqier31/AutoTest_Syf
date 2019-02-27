#coding:utf-8

from wordcloud import WordCloud,ImageColorGenerator
import matplotlib.pyplot as plt
from scipy.misc import imread
import jieba


#读取一个txt文件

text = open(r'C:\temp/bbb.txt','r').read()

#读入背景图片
bg_pic = imread(r'C:\temp/1.png')

wordlist_after_jieba = jieba.cut(text, cut_all = True)
wl_space_split = " ".join(wordlist_after_jieba)

#生成词云
font = r'C:\Windows\Fonts\simfang.ttf'
wc = WordCloud(mask=bg_pic,background_color='white',font_path=font,    scale=1.5).generate(wl_space_split)
image_colors = ImageColorGenerator(bg_pic)
#显示词云图片

plt.imshow(wc)
plt.axis('off')
plt.show()
# 保存图片
wc.to_file(r'C:\temp/66.jpg')
