1# Java Analysis Excel Tool easyexcel
Most well-known frameworks for Java parsing and generating Excel files are Apache POI and JXL. But they all have a serious problem, they are very memory-intensive. POI has a set of SAX mode API that can solve some memory overflow problems to a certain extent, but POI still has some defects, such as the decompression and decompression storage for Excel 2007, done in memory causing the memory consumption to be very high. 

Easyexcel rewrites POI's analysis of Excel 2007. It can still use a POI sax for 3M excel. It still needs about 100M memory to be reduced to KB level, and even with large excel files it will not have memory overflows. Excel 2003 version relies on POI sax mode. The model conversion package is build on top to make the usage more convenient.

## Related documents
* [About](/abouteasyexcel.md)
* [Quickstart](/quickstart.md)
* [Problems](/problem.md)
* [Update](/update.md)
* [Original README](/README.md)

## Installation 

<dependency>
    <groupId>com.alibaba</groupId>
    <artifactId>easyexcel</artifactId>
    <version>{latestVersion}</version>
</dependency>

## Latest version：1.1.2-beta4
## Maintainer
Jīpéngfēi/姬朋飞（玉霄/Yù xiāo）
## Quick start
### Reading an Excel file
Test code：[https://github.com/alibaba/easyexcel/blob/master/src/test/java/com/alibaba/easyexcel/test/ReadTest.java](/src/test/java/com/alibaba/easyexcel/test/ReadTest.java)

Excel 2007: read less than 1000 rows of data and return List<List<String>>
```
List<Object> data = EasyExcelFactory.read(inputStream, new Sheet(1, 0));
```
Excel 2007: read less than 1000 rows of data and return List<? extend BaseRowModel>
```
List<Object> data = EasyExcelFactory.read(inputStream, new Sheet(2, 1,JavaModel.class));
```
Excel 2007: read more than 1000 rows of data and return List<List<String>>
```
ExcelListener excelListener = new ExcelListener();
EasyExcelFactory.readBySax(inputStream, new Sheet(1, 1), excelListener);
```

Excel 2007: read more than 1000 rows of data and return List<? extend BaseRowModel>
```
ExcelListener excelListener = new ExcelListener();
EasyExcelFactory.readBySax(inputStream, new Sheet(2, 1,JavaModel.class), excelListener);
```
Excel 2003 (same as above)

### Writing an Excel
测试代码地址：[https://github.com/alibaba/easyexcel/blob/master/src/test/java/com/alibaba/easyexcel/test/WriteTest.java](/src/test/java/com/alibaba/easyexcel/test/WriteTest.java)
没有模板
```OutputStream out = new FileOutputStream("/Users/jipengfei/2007.xlsx");
ExcelWriter writer = EasyExcelFactory.getWriter(out);

//写第一个sheet, sheet1  数据全是List<String> 无模型映射关系
Sheet sheet1 = new Sheet(1, 3);
sheet1.setSheetName("第一个sheet");
//设置列宽 设置每列的宽度
Map columnWidth = new HashMap();
columnWidth.put(0,10000);columnWidth.put(1,40000);columnWidth.put(2,10000);columnWidth.put(3,10000);
sheet1.setColumnWidthMap(columnWidth);
sheet1.setHead(createTestListStringHead());
//or 设置自适应宽度
//sheet1.setAutoWidth(Boolean.TRUE);
writer.write1(createTestListObject(), sheet1);

//写第二个sheet sheet2  模型上打有表头的注解，合并单元格
Sheet sheet2 = new Sheet(2, 3, JavaModel1.class, "第二个sheet", null);
sheet2.setTableStyle(createTableStyle());
writer.write(createTestListJavaMode(), sheet2);

//写第三个sheet包含多个table情况
Sheet sheet3 = new Sheet(3, 0);
sheet3.setSheetName("第三个sheet");
Table table1 = new Table(1);
table1.setHead(createTestListStringHead());
writer.write1(createTestListObject(), sheet3, table1);

//写sheet2  模型上打有表头的注解
Table table2 = new Table(2);
table2.setTableStyle(createTableStyle());
table2.setClazz(JavaModel1.class);
writer.write(createTestListJavaMode(), sheet3, table2);

//关闭资源
writer.finish();
out.close();
```
有模板
```InputStream inputStream = new BufferedInputStream(new FileInputStream("/Users/jipengfei/temp.xlsx"));
OutputStream out = new FileOutputStream("/Users/jipengfei/2007.xlsx");
ExcelWriter writer = EasyExcelFactory.getWriterWithTemp(inputStream,out,ExcelTypeEnum.XLSX,true);

//写第一个sheet, sheet1  数据全是List<String> 无模型映射关系
Sheet sheet1 = new Sheet(1, 3);
sheet1.setSheetName("第一个sheet");
//设置列宽 设置每列的宽度
Map columnWidth = new HashMap();
columnWidth.put(0,10000);columnWidth.put(1,40000);columnWidth.put(2,10000);columnWidth.put(3,10000);
sheet1.setColumnWidthMap(columnWidth);
sheet1.setHead(createTestListStringHead());
//or 设置自适应宽度
//sheet1.setAutoWidth(Boolean.TRUE);
writer.write1(createTestListObject(), sheet1);

//写第二个sheet sheet2  模型上打有表头的注解，合并单元格
Sheet sheet2 = new Sheet(2, 3, JavaModel1.class, "第二个sheet", null);
sheet2.setTableStyle(createTableStyle());
writer.write(createTestListJavaMode(), sheet2);

//写第三个sheet包含多个table情况
Sheet sheet3 = new Sheet(3, 0);
sheet3.setSheetName("第三个sheet");
Table table1 = new Table(1);
table1.setHead(createTestListStringHead());
writer.write1(createTestListObject(), sheet3, table1);

//写sheet2  模型上打有表头的注解
Table table2 = new Table(2);
table2.setTableStyle(createTableStyle());
table2.setClazz(JavaModel1.class);
writer.write(createTestListJavaMode(), sheet3, table2);

//关闭资源
writer.finish();
out.close();
```

### web下载实例写法
```
public class Down {
    @GetMapping("/a.htm")
    public void cooperation(HttpServletRequest request, HttpServletResponse response) {
        ServletOutputStream out = response.getOutputStream();
        ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX, true);
        String fileName = new String(("UserInfo " + new SimpleDateFormat("yyyy-MM-dd").format(new Date()))
                .getBytes(), "UTF-8");
        Sheet sheet1 = new Sheet(1, 0);
        sheet1.setSheetName("第一个sheet");
        writer.write0(getListString(), sheet1);
        writer.finish();
        response.setContentType("multipart/form-data");
        response.setCharacterEncoding("utf-8");
        response.setHeader("Content-disposition", "attachment;filename="+fileName+".xlsx");
        out.flush();
        }
    }
}
```
### 联系我们
有问题阿里同事可以通过钉钉找到我，阿里外同学可以通过git留言。其他技术非技术相关的也欢迎一起探讨。
### 招聘&交流
阿里巴巴新零售事业部--诚招JAVA资深开发、技术专家。有意向可以微信联系，简历可以发我邮箱jipengfei.jpf@alibaba-inc.com
或者加微信：18042000709

<img src="https://github.com/alibaba/easyexcel/blob/master/img/WechatIMG8.png" width="30%" height="30%" />
