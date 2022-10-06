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

from PIL import Image, ImageFont
from handright import Template, handwrite
import docx


# 自动缩进排版，如果已在word里设置缩进可以注释本段
# indent_size控制缩进,file_path文档路径
def get_text(file_path, indent_size=4):
    doc = docx.Document(file_path)
    texts = []
    indent = ''
    for i in range(0, indent_size):
        indent = indent + ' '
    for paragraph in doc.paragraphs:
        texts.append(indent + paragraph.text)
        # print(texts)

    return '\n'.join(texts)


# 根目录下的word文档
text = get_text('1.docx')
print(text)
template = Template(
    background=Image.new(mode="1", size=(1750, 2479), color=1),  # 自定义背景图片

    font=ImageFont.truetype("fonts\\李国夫手写体.ttf",size=100),  # 字体选择手写体
    line_spacing=120,
    # fill=(0, 0, 0),  # 字体颜色，括号内为RGB的值
    left_margin=100,
    top_margin=100,
    right_margin=100,
    bottom_margin=100,
    word_spacing=15,
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

