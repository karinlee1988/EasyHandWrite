#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time : 2023/1/8 21:13
# @Author : karinlee
# @FileName : easy_handwrite.py
# @Software : PyCharm
# @Blog : https://blog.csdn.net/weixin_43972976
# @github : https://github.com/karinlee1988/
# @gitee : https://gitee.com/karinlee/
# @Personal website : https://karinlee.cn/


"""
本库通过读取word文档的内容，转化为手写体图片。
生成的图片使用合适的打印机打印在活页纸上再装订成笔记本，
跟真实手写的笔记不能说完全相同，只能说一模一样=。=

以下是图片像素设置参考：

B5 SIZE 300PPI: (2079, 2953)

A5 SIZE 300PPI: (1750, 2479)

20230108 tests OK

"""

from PIL import Image, ImageFont
from handright import Template, handwrite
import docx

class EasyHandWrite(object):
    """
    将word文档内容生成手写体图片

    20230108 tests OK

    """
    def __init__(self,wordfile):
        self.wordfile = wordfile
        # self.font = font

    def get_text(self,indent_size=4):
        # 自动缩进排版，如果已在word里设置缩进可以注释本段
        # indent_size控制缩进,file_path文档路径
        doc = docx.Document(self.wordfile)
        texts = []
        indent = ''
        for i in range(0, indent_size):
            indent = indent + ' '
        for paragraph in doc.paragraphs:
            texts.append(indent + paragraph.text)
            # print(texts)
        return '\n'.join(texts)

    def main(self):
        text = self.get_text()
        # print(text)
        template = Template(
            # 这里设置生成图片的尺寸。size参数可参考最上方注释处，现在是A5 size，根据实际情况自行修改
            background=Image.new(mode="1", size=(1750, 2479), color=1),  # 自定义背景图片
            font=ImageFont.truetype("fonts\\李国夫手写体.ttf",size=85),  # 字体选择手写体
            line_spacing=110,
            # fill=(0, 0, 0),  # 字体颜色，括号内为RGB的值
            left_margin=180,   # 左页面边距
            top_margin=300,    # 上页面边距
            right_margin=180,  # 右页面边距
            bottom_margin=500, # 底部页面边距
            word_spacing=7,    # 字间距
            line_spacing_sigma=2,  # 行间距随机扰动
            font_size_sigma=3,  # 字体大小随机扰动
            word_spacing_sigma=1,  # 字间距随机扰动
            end_chars="，。",  # 防止特定字符因排版算法的自动换行而出现在行首
            perturb_x_sigma=2,  # 笔画横向偏移随机扰动
            perturb_y_sigma=2,  # 笔画纵向偏移随机扰动
            perturb_theta_sigma=0.05,  # 笔画旋转偏移随机扰动
        )
        images = handwrite(text, template)
        for i, im in enumerate(images):
            print(i,im)
            assert isinstance(im, Image.Image)
            # im.show()
            im.save(f'pic\\{i}.png')

if __name__ == '__main__':
    # 读取SOURCE.docx的内容，生成的手写体图片存放在该项目pic文件夹下，
    # 根据内容的多少生成0.png，1.png，2.png...，可以后续将图片另行转换为pdf文档再打印。
    write = EasyHandWrite(wordfile='SOURCE.docx')
    write.main()