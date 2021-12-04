# FastExcel

A Java lib to read excel providing stream-processing program model. It is very suitable for online/web applications where requires
real-time and being sensitive to resource consumptions like CPU and memory.

## Features

* Stream-processing program model;
* Faster read and less memory consumption;

## Supported File Formats
1. 07 version excel files like xlsx, xlsm, etc;
2. csv files

# Inspired By

1. [EasyExcel](https://github.com/alibaba/easyexcel)
2. [Xlsx Streamer](https://github.com/monitorjbl/excel-streaming-reader)

# Usage Example

## Xlsx Read

```java
public static void main(String[]args){
        InputStream inputStream = new FileInputStream("src/example.xlsx");
        XlsxStream stream = XlsxStreamBuilder.builder()
                                .stream(this.inputStream)
                                .build();
        
        stream.foreach(row -> doSothing(row));
}
```

# Contribution

## Set up

* At project root path, run command 
```shell
./gradlew idea
```

* Then open ``fastexcel.ipr`` by IntelliJ

## Compile Project
* At project root path, run command
```shell
./gradlew build
```