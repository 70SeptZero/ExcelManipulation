# ExcelManipulation

# 语言Language

中文说明在前半段。

The English instructions are in the second half, by Translate tools and me.

# 项目介绍
    本项目是用于合并Excel文件而写的一个项目。
    能够将多个Excel中的所有表（除隐藏表），合并到一个Sheet中，并且会自动清除无内容的行。
    
    将需要合并的Excel文件放置在excel命名文件夹，运行项目后，即可在output中获取到合并结果，输出的命名格式为：Manipulation_时间戳.xlsx。
    
    相对路径布局如下：
    ├── excel                   //放需要合并的文件
    ├── output                  //输出文件夹
    ├── ExcelManipulation       //项目
    
    这两个文件夹是放在和项目同一级下面的，不是在项目里面。
    
    未来更新企划：
    如果不知道如何启动的，后续会再做一个简单版本，点击即可启动。
    之后再单独做成一个方法，可以引入后直接调用。
    
    若有遇到什么问题，后续记录更新。

# 环境依赖

jdk-17.0.9


# 目录结构描述
    ├── ReadMe.md           // 帮助文档
    ├── src                 // 代码
    │   ├── main
    │       └── java.com.septzero.excelmanipulation
    │       	└── ExcelManipulationApplication.java    //代码
    │   └── test
    │       └── java.com.septzero.excelmanipulation
    │       	└── ExcelManipulationApplicationTests.java
    ├── .gitignore
    ├── LICENSE
    └── pom.xml

# 使用说明

1、将需要合并的Excel文件放入excel文件夹中

2、运行项目

3、在output中获取合并后的文件

注：由于poi只支持一部分excel公式，所以在遇到公式的时候，可能出现无法正常获取公式类单元格的值。

遇到这种情况的时候，项目会将该格子的来源和输出文件的位置打印出来。

输出文件中会写入该公式，而没有值。

# 版本内容更新
## v0.0.97: 
    1、实现所有表数据的合并
    2、排除了隐藏表，隐藏表将不参与合并

##  v0.0.98: 

```
1、排除了带格式的空白行
```

##  v0.0.9: 

```
1、解决了带公式单元格合并异常的问题
2、增加了对不支持公式的处理
```

##  v1.0.0: 

```
1、上传至gitHub
2、实现了将除隐藏表外的所有表按顺序合并至一张表中
3、排除了各类空行、带格式空行的影响
4、支持对公式类单元格的处理，增加了对不支持公式的处理
```

 

------

# English

# Project Introduction

```
This project is designed to merge Excel files.

It is able to merge all sheets(excluding hidden sheets) from multiple Excel files into one Sheet, and automatically removing empty rows.

Put the target Excel files in the folder named "excel". After running the project, you can get the merge result in the folder named "output", and the output file will be named as "Manipulation_ timestamp.xlsx".

The relative path structure is as follows:
├── excel             // folder for Excel files to be merged
├── output            // output folder
├── ExcelManipulation // project folder

These two folders are placed at the same level as the project, not within the project.

Future Renewal Plans:
If you don't know how to launch, there will be a simple version in the future, click to launch.
After that, it can be made into a separate method, which can be directly called after importing.

If there are any problems or suggestions, I will deal with them in the feature.
```

# Environment

jdk-17.0.9

# Directory structure

```
├── ReadMe.md           // Help Documentation
├── src                 // code
│   ├── main
│       └── java.com.septzero.excelmanipulation
│       	└── ExcelManipulationApplication.java    //code
│   └── test
│       └── java.com.septzero.excelmanipulation
│       	└── ExcelManipulationApplicationTests.java
├── .gitignore
├── LICENSE
└── pom.xml
```

# Directions for use

1、Place the Excel files to be merged into the "excel" folder.

2、Run the project.

3、Retrieve the merged file from the "output" folder.

Note: Due to limited support for some Excel formulas in Apache POI, it may not be possible to retrieve the values of cells containing certain formulas.

When this situation occurs, the project will print the source of the cell and the location in the output file. The output file will contain the formula without its calculated value.

# Version

## v0.0.97: 

```
1、Implemented merging of all sheet data.
2、Excluded hidden sheets from the merging process.
```

##  v0.0.98: 

```
1、Excluded formatted empty rows.
```

##  v0.0.9: 

```
1、Resolved the issue of abnormal merging of cells with formulas.
2、Added handling for unsupported formulas.
```

##  v1.0.0: 

```
1、Uploaded to GitHub.
2、Implemented merging of all sheets, excluding hidden sheets, into one sheet in sequential order.
3、Excluded various types of empty rows and formatted empty rows.
4、Supported handling of cells with formulas and added handling for unsupported formulas.
```