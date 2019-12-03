# 使用Python在Excel里面画画
## 起因
之前看到过很多头条，说哪国某人坚持了多少年自学使用excel画画，效果十分惊艳。对于他们的耐心我十分敬佩，但是作为一个程序员，自然也得挑战一下自己。这种需求，我们十分钟就可以完成！
## 运行环境
* [Python 3.7](https://www.python.org/downloads/release/python-370/)  
* [PIL 5.3](https://pypi.org/project/Pillow/)
* [openpyxl 2.5.12](https://pypi.org/project/openpyxl/)
> `pip install openpyxl==2.5.12`  
>  `pip install Pillow==5.3.0`

## 最终效果  
![monalisa](https://raw.githubusercontent.com/alisen39/drawExcel/master/image-20191117175137916.png)

## 鸣谢  
* 感谢掘友[itbj00](https://juejin.im/user/5ad1c374f265da23a4053f62)指出openpyxl 2.6.2时的版本兼容问题。  
``` python
# 当openpyxl版本大于2.6.1时需修改 column_dimensions 索引为字母形式

# openpyxl <= 2.5.12   
worksheet.column_dimensions[_w].width = 1

# openpyxl >= 2.6.1
_w_letter = openpyxl.utils.get_column_letter(_w)  
worksheet.column_dimensions[_w_letter].width = 1
```

