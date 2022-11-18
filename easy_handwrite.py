#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time : 2022/10/7 0:18
# @Author : karinlee
# @FileName : easy_handwrite.py
# @Software : PyCharm
# @Blog : https://blog.csdn.net/weixin_43972976
# @github : https://github.com/karinlee1988/
# @gitee : https://gitee.com/karinlee/
# @Personal website : https://karinlee.cn/


"""
B5 SIZE 300PPI: (2079, 2953)

A5 SIZE 300PPI: (1750, 2479)

"""
import sys
from PIL import Image, ImageFont
from handright import Template, handwrite
import docx

class EasyHandWrite(object):
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
            background=Image.new(mode="1", size=(1750, 2479), color=1),  # 自定义背景图片
            font=ImageFont.truetype("fonts\\加林手写.ttf",size=85),  # 字体选择手写体
            line_spacing=110,
            # fill=(0, 0, 0),  # 字体颜色，括号内为RGB的值
            left_margin=180,
            top_margin=300,
            right_margin=180,
            bottom_margin=500,
            word_spacing=7,
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
            im.save(f'{i}.png')

if __name__ == '__main__':
    write = EasyHandWrite(wordfile='SOURCE.docx')
    write.main()