## DocxGenerator


### 简介

一个根据模板 docx 文档和数据表单快速生成多个成品 docx 文档的小工具.

作为某某活动打工人的的你, 是否会偶尔接到给每个活动成员制作一些需要写上名字、编号等个性化信息但是大部分东西又是一样的玩意。

你可能手上有一份名单和模板, 但是, 你真的要手工一个个改吗, 是时候让 DocxGenerator 解放你的双手了！

本项目采用 **MIT License** 开源.

### 使用方式

本项目 example 文件夹中有示例和效果演示.

点击 [Release](https://github.com/huan-yp/DocxGenerator/releases) 找到本项目最新发布并下载解压.

选择解压到 `convert/` 目录, 点击 .exe 文件运行, **不要删去剩下的任何非 .exe 文件**

#### 制作模板 .docx 文档

下面以制作准考证 .docx 文件为例:

1. 用第一份考生信息手工做好一个模板 template.docx 文件.

2. 将每一个需要替换的信息改为 `PLACE_HOLDER_[数字ID]`.
    - 数字ID **不能含有前导零、必须是从 1 开始的连续数字**。
    - 输入 数字ID 时不能有中断, (以第 2 个需要替换的信息为例)我建议你现在记事本上输入好 `PLACE_HOLDER_2`, 用 `ctrl+c` 复制一份, 在 .docx 文档里选中需要被替换的内容, 再 `ctrl+v` 粘贴上去, 这样的输入一定不会存在中断.

3. 保存经过 2 修改的文件 

#### 制作数据表

1. 数据表是一个 xlsx 表格(excel), 第一行必须是表头, 从第一列开始, 内容均为 `PLACE_HOLDER_[数字ID]`, 数字ID 规则同上.

2. 每一行为一份数据, 为每一份数据的每一个 HOLDER 填写对应的替换值.

3. 保存完成 2 的文件

#### 填写 config.yaml

如果你的计算机知识不够扎实, 所有路径请使用**绝对路径**.

如果你想要使用**相对路径**, 请注意，使用相对 `config.yaml` 的路径.

本项目的 release 用 pyinstaller 打包, 拥有相关知识的同学可以自行替换为相对路径, 实际上我更建议你下载本项目源代码使用.

给出一个示例的文件结构和配置文件如下

```
E:/IPP/DocxGenerator/example/
├── src/
│   ├── config.yaml
│   ├── template.docx
│   └── data.xlsx
└── result/
```

```
template: "E:/IPP/DocxGenerator/example/src/template.docx" # 模板 docx 文件路径
data: "data.xlsx" # 数据 xlsx 表格路径
dst: "../result" # 目标路径, 如果不存在, 将会创建
filename: # 命名方式, 留空直接采用 [对应数据的行号 - 1] 命名
```

- `template`: 这里使用了绝对路径.
- `data`: 相对 `config.yaml` 的路径.
- `dst`: 相对 `config.yaml` 的路径，支持向上一级目录的索引.
- `filename`: 它支持 `PLACE_HOLDER_[数字ID]` 参数, 会对次做替换, 为了避免重名, 在此基础上会加上 `[对应数据的行号 - 1]-` 作为前置

#### 运行

- 一般来说双击 Release 的 .exe 文件运行即可, 如果发现运行出错, 可以用命令行运行它, 然后将报错信息提交到 Issue

- 如果你对自己书写的绝对路径不够自信, 尝试拖拽你的 `config.yaml` 文件到程序运行窗口, 它会自动帮你填写路径

### Release


### Developer Use

如果你了解 Python 相关知识, 那么你可以考虑直接 clone 本项目并通过源代码调用.

### Contribute

- 欢迎在 Issues 中反馈 bug、请求新功能
- 欢迎提交 PR

### Bonus

- 动动手指点击右上角的 Star 是对作者最好的支持与鼓励.