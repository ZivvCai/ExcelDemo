package com.czw.testDemo;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.Collection;

public class WriteExcel {

    //针对XLSX
    private File file;
    private OutputStream os;
    private Workbook workbook = null;

    //构造函数
    public WriteExcel(File file) throws InvalidFormatException,IOException,ClassNotFoundException {
        this.file = file;
        if(!file.exists())
            file.createNewFile();
        os = new FileOutputStream(file);
        workbook = new XSSFWorkbook();//如果是XLS格式则改成HSSFWorkbook
        Sheet sheet = workbook.createSheet("User");
        Field[] fields = Class.forName("com.czw.testDemo.User").getDeclaredFields();
        Row titleRow = sheet.createRow(0);
        for(int i = 0; i<fields.length; i++){
            Cell cell = titleRow.createCell(i);
            cell.setCellValue(fields[i].getName());
        }
    }
    //将单个Bean对象写入Excel
    public void write(User user) throws IOException{
        Sheet sheet = workbook.getSheet("User");
        int lastRowNum = sheet.getLastRowNum();
        Row index = sheet.createRow(lastRowNum + 1);
        //通过反射机制获取类的属性以及get方法（通用）
        Class clazz = user.getClass();//Class.forName("Bean类名");//Bean.class
        Field[] fields = clazz.getDeclaredFields();
//        index.createCell(0).setCellFormula("ROW() - 1");
        for(int i = 0; i<fields.length; i++){
            try {
                index.createCell(i).setCellValue((String)clazz.getMethod("get"+toUpperFirstWord(fields[i].getName()),
                        null).invoke(user,null));
            }catch (Exception e){
                e.printStackTrace();
            }

        }
    }
    //将多个Bean对象写入Excel
    public void write(Collection<User> users) throws IOException{
        for(User user : users){
            this.write(user);
        }
    }
    // 创建文件输出流，准备输出电子表格：这个方法发在调用前面的方法之后必须执行，否则你在sheet上做的任何操作都不会有效
    public void execute() throws IOException{
        workbook.write(os);
        workbook.close();
    }

    //将字符串的首字母大写
    private String toUpperFirstWord(String str){
        if(str != null && str != ""){
            str  = str.substring(0,1).toUpperCase()+str.substring(1);
        }
        return str;
    }

}
