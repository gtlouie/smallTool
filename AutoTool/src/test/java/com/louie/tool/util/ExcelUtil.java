package com.louie.tool.util;


import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class ExcelUtil {

    //预期数据文件
    private static final String FILE_EXPECT="/Users/gttester/Documents/workScript/AutoTool/src/test/resources/BA数据需求.xlsx";
    //读取Excel的哪个Sheet
    private static int  SHEET_EXPECT=0;
    //需要对比的列数
    private static int  COL_EXPECT=0;


    //实际数据文件
    private static final String FILE_ACTUAL = "/Users/gttester/Documents/workScript/AutoTool/src/test/resources/前端数据.xlsx";
    //读取Excel的哪个Sheet
    private static int  SHEET_ACTUAL=0;
    //需要对比的列数
    private static int  COL_ACTUAL=0;


    //对比数据结果文件
    private static final String FILE_RESULT = "/Users/gttester/Documents/workScript/AutoTool/src/test/resources/对比结果.xlsx";
    //对比数据结果文件第一行表头
//    private static final String[] RESULT_TITLE={"BA需求英文版-类别","FE英文版-类别","对比结果"};
    private static final String[] RESULT_TITLE={"BA","FE","对比结果"};
    //对比结果描述
    private  static final String RESULT_PASS="测试通过，对比数据一致";
    private  static final String RESULT_FAIL="测试不通过，对比数据一致";
    private  static  int SUM_PASS=0;
    private  static  int SUM_FAIL=0;




    //获取Excel sheet 操作对象
    public static Sheet getSheet(String fileName,int SheetNum) {

        try{
            //获取文件
            File file = new File(fileName);
            //获取输入流
            InputStream stream = new FileInputStream(file);
            //打开文件
            Workbook xssfWorkbook = new XSSFWorkbook(stream);

            //读取excel的第一个Sheet
            Sheet sheet = xssfWorkbook.getSheetAt(SheetNum);

            return sheet;

        }catch (Exception e){
            System.out.println(e);
        }
        return null;
    }

    //对比两个sheet 的数据的准确性
    public static void compareExcelData(Sheet expectSheet,int colExpect,Sheet actualSheet,int colActual){
        int rowExpTotals=expectSheet.getPhysicalNumberOfRows();
        int rowActTotals=actualSheet.getPhysicalNumberOfRows();
        //存储预期数据
        List<String> expectList=new ArrayList<String>();
        //存储实际数据
        List<String> actualList=new ArrayList<String>();
        //存储对比结果
        List<String> resultList=new ArrayList<String>();
        for(int i=0;i<rowExpTotals-1;i++){
            //读取第i行
            Row rowExpect = expectSheet.getRow(i);
            Row rowActual=actualSheet.getRow(i);
            //读取第一行的第colExpect列
            String valueExpect=rowExpect.getCell(colExpect).getStringCellValue();
            String valueActual=rowActual.getCell(colActual).getStringCellValue();
            //输出读取第一行的第一列数据
            expectList.add(valueExpect);
            actualList.add(valueActual);

           // 对比数据
            if(valueActual.equals(valueExpect)){
                resultList.add(RESULT_PASS);
                SUM_PASS++;
                System.out.println("测试通过"+"实际数据："+valueActual+"  = "+"预期结果："+valueExpect);

            }else{
                resultList.add(RESULT_FAIL);
                SUM_FAIL++;
                System.out.println("测试不通过"+"实际数据："+valueActual+" != "+"预期结果："+valueExpect);

            }

        }
        //将数据写入测试结果
        ExcelUtil.writeExcel(FILE_RESULT,RESULT_TITLE,expectList,actualList,resultList);
        System.out.println("测试结果：(对比总数据:"+expectList.size()+" 测试通过条数:"+SUM_PASS+" 测试不通过条数:"+SUM_FAIL+" )");

    }

    //将预期数据、实际数据、对比结果数据写入测试结果中
    public static void writeExcel(String filename,String[] title,List<String> expectList,List<String> actualList,List<String> resultList){
        try {
                //获取文件
                File file = new File(filename);
                //获取输入流
                OutputStream outputStream = new FileOutputStream(file);
                //创建工作簿
                XSSFWorkbook xssfWorkbook = null;
                xssfWorkbook = new XSSFWorkbook();
                //创建工作表
                XSSFSheet xssfSheet;
                xssfSheet = xssfWorkbook.createSheet();
                //创建行
                XSSFRow xssfRow;
                //创建列，即单元格Cell
                XSSFCell xssfCell;
                //从第1行开始写入
                xssfRow = xssfSheet.createRow(0);
                //写入标题
                for(int x=0;x<title.length;x++){
                    xssfRow.createCell(x).setCellValue(title[x]);
                }
                //把List里面的数据写到excel中
                for (int i = 1; i < expectList.size(); i++) {
                    //从第2行开始写入
                    xssfRow = xssfSheet.createRow(i);
                    xssfRow.createCell(0).setCellValue(expectList.get(i));
                    xssfRow.createCell(1).setCellValue(actualList.get(i));
                    xssfRow.createCell(2).setCellValue(resultList.get(i));

                }

                //用输出流写到excel
                try {
                    xssfWorkbook.write(outputStream);
                    outputStream.flush();
                    outputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
        }catch (Exception e){
            System.out.println(e);
        }

    }

    public static void main(String[] args) throws Exception{

          //获得预期Excel中的数据
           Sheet sheetExpect=ExcelUtil.getSheet(FILE_EXPECT,SHEET_EXPECT);

           //获取实际Excel 中的数据
           Sheet sheetActual=ExcelUtil.getSheet(FILE_ACTUAL,SHEET_ACTUAL);

           //预期sheet 的第n列数据 对比 实际sheet 的第m 列数据
           ExcelUtil.compareExcelData(sheetExpect,COL_EXPECT,sheetActual,COL_ACTUAL);


    }


}
