参照接口、实现属性、实现函数
参考：.net python ...

 ```
 _     _                               _            _
| |   (_)_ __   ___  __ _ _ __    __ _| | __ _  ___| |__  _ __ __ _
| |   | | '_ \ / _ \/ _` | '__|  / _` | |/ _` |/ _ \ '_ \| '__/ _` |
| |___| | | | |  __/ (_| | |    | (_| | | (_| |  __/ |_) | | | (_| |
|_____|_|_| |_|\___|\__,_|_|     \__,_|_|\__, |\___|_.__/|_|  \__,_|
                                         |___/
```

强大、安全、高可读的 VBS 线性代数类库。含向量、矩阵等实现。

使用不可变类，以提高安全性。

# 浏览介绍

- [中文](Readme_zh.md)
- [英文](Readme.md)

# 开始

## 环境要求

- 视窗操作系统（XP SP2 或更高版本）

## 安装

以**管理员权限**运行以下命令：

```
git clone https://github.com/OldLiu001/LinearAlgebra.vbs.git
cd LinearAlgebra.vbs
regsvr32 LinearAlgebra.wsc
```

**警告：不要使用右键菜单注册 *LinearAlgebra.wsc* 。**

## 用法

```
Set objVectorGenerator = CreateObject("LinearAlgebra.VectorGenerator")
Set objMatrixGenerator = CreateObject("LinearAlgebra.MatrixGenerator")
```

|类/对象|属性/方法|名称|参数|返回值|简介|
|:---:|:---:|:---:|:---:|:---:|:---:|
|VectorGenerator|方法|Init|一维数组|Vector 对象|创建向量|
|VectorGenerator|方法|Zero|维度|Vector 对象|创建零向量|
|MatrixGenerator|方法|Init|二维数组/Vector 对象|Matrix 对象|创建矩阵|
|MatrixGenerator|方法|Zero|行数、列数|Matrix 对象|创建零矩阵|
|MatrixGenerator|方法|Identity|维度|Matrix 对象|创建单位矩阵|


更多示例：[/Examples](Examples/)

## 错误处理