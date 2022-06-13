package com.example.demoexcel.easyexcel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.read.builder.ExcelReaderBuilder;
import com.alibaba.excel.read.builder.ExcelReaderSheetBuilder;
import com.alibaba.excel.support.ExcelTypeEnum;
import lombok.SneakyThrows;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public class EasyExcelDemo {

    @SneakyThrows
    public static void main(String[] args) {
//        ExcelReaderBuilder s = EasyExcel.read(new File("D:\\Document\\WeChat Files\\wxid_1118631186712\\FileStorage\\MsgAttach\\c804d80638890b4f4284299f5460c3d1\\File\\2022-06\\11.xlsx"));
//
//        ExcelReaderSheetBuilder ss = EasyExcel.readSheet();
//        String name = ss.build().getSheetName();
//        System.out.println("name=="+name);
//        InputStream inputStream = new FileInputStream("D:\\Document\\WeChat Files\\wxid_1118631186712\\FileStorage\\MsgAttach\\c804d80638890b4f4284299f5460c3d1\\File\\2022-06\\11.xlsx");
//        ExcelListener listener = new ExcelListener();
//        ExcelReader excelReader = new ExcelReader(inputStream, ExcelTypeEnum.XLS, null, listener);
//        excelReader.read();


        // 读取 excel 表格的路径
        String readPath = "D:\\Document\\WeChat Files\\wxid_1118631186712\\FileStorage\\MsgAttach\\c804d80638890b4f4284299f5460c3d1\\File\\2022-06\\11.xlsx";

//        try {
//            // sheetNo --> 读取哪个 表单
//            // headLineMun --> 从哪一行开始读取( 不包括定义的这一行，好比 headLineMun为2 ，那么取出来的数据是从 第三行的数据开始读取 )
//            // clazz --> 将读取的数据，转化成对应的实体，须要 extends BaseRowModel
//            Sheet sheet = new Sheet(1, 1);
//
//            // 这里 取出来的是 ExcelModel实体 的集合
////            List<Object> readList = EasyExcelFactory.read(new FileInputStream(readPath), sheet);
////            // 存 ExcelMode 实体的 集合
////            List<Object> list = new ArrayList<Object>();
////            for (Object obj : readList) {
////                list.add(obj+"222222222222");
////            }
//
//            // 取出数据
//            StringBuilder str = new StringBuilder();
//            str.append("{");
//            String link = "";
////            for (Object mode : list) {
////                str.append(link).append("\""+mode.getColumn1()+"\":").append("\""+mode.getColumn2()+"\"");
////                link= ",";
////            }
////            str.append("};");
//            System.out.println(str);
//
//        } catch (FileNotFoundException e) {
//            e.printStackTrace();
//        }

    }
}
