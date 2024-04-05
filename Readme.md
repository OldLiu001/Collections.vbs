施工中！！

本类库 < com对象 < net帮助
本类库全.net兼容

 ```
 _     _                               _            _
| |   (_)_ __   ___  __ _ _ __    __ _| | __ _  ___| |__  _ __ __ _
| |   | | '_ \ / _ \/ _` | '__|  / _` | |/ _` |/ _ \ '_ \| '__/ _` |
| |___| | | | |  __/ (_| | |    | (_| | | (_| |  __/ |_) | | | (_| |
|_____|_|_| |_|\___|\__,_|_|     \__,_|_|\__, |\___|_.__/|_|  \__,_|
                                         |___/
```

A powerful, safe, readable, easy to use Visual Basic Script Library for Linear Algebra Operations. Including Vector, Matrix, and so on.

Use immutable classes for safety.

# View introduction in

- [Chinese](Readme_zh.md)
- [English](Readme.md)

# Getting Started

## Requirements

- A Windows OS (Windows XP SP2 or later)

## Installation

Run following commands as **administrator**:

```
git clone https://github.com/OldLiu001/LinearAlgebra.vbs.git
cd LinearAlgebra.vbs
regsvr32 LinearAlgebra.wsc
```

**WARN: DO NOT REGISTER *LinearAlgebra.wsc* BY RIGHT CLICKING ON IT.**

## Usage

```
Set objVectorGenerator = CreateObject("LinearAlgebra.VectorGenerator")
Set objMatrixGenerator = CreateObject("LinearAlgebra.MatrixGenerator")
```

|Class/Object|Property/Method|Name|Argument(s)|Return Value|Description|
|:---:|:---:|:---:|:---:|:---:|:---:|
|VectorGenerator|Method|Init|Array1D|Vector Object|Generate a Vector|
|VectorGenerator|Method|Zero|Dimension|Vector Object|Generate a Zero Vector|
|MatrixGenerator|Method|Init|Array2D/Vector|Matrix Object|Generate a Matrix|
|MatrixGenerator|Method|Zero|Row, Column|Matrix Object|Generate a Zero Matrix|
|MatrixGenerator|Method|Identity|Dimension|Matrix Object|Generate an Identity Matrix|


See more Examples in [/Examples](Examples/).

## Error Handling
