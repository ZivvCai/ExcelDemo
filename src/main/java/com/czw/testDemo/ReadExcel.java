package com.czw.testDemo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ReadExcel {

    //读取XLSX文档
    public List<User> readXLSX(File file) throws InvalidFormatException, IOException{
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = new XSSFWorkbook(file).getSheetAt(0);//也可以通过Sheet名获取Sheet，getSheet("name")
        List<User> result = new ArrayList<User>();
        int rowStart = sheet.getFirstRowNum() + 1;
        int rowEnd = sheet.getLastRowNum();

        for(int i = rowStart; i < rowEnd + 1; i++){
            Row row = sheet.getRow(i);//获取sheet的行对象
            User user = this.getBeanFromRow(row);
            if(user != null)
                result.add(user);
        }
        return result;
    }
    //读取XLS文档
    public List<User> readXLS(File file) throws InvalidFormatException, IOException{
        POIFSFileSystem poifsFileSystem = new POIFSFileSystem(new FileInputStream(file));
        Sheet sheet = new HSSFWorkbook(poifsFileSystem).getSheetAt(0);
        List<User> result = new ArrayList<User>();
        int rowStart = sheet.getFirstRowNum() + 1;
        int rowEnd = sheet.getLastRowNum();

        for(int i = rowStart; i < rowEnd + 1; i++){
            Row row = sheet.getRow(i);
            User user = this.getBeanFromRow(row);
            if(user != null)
                result.add(user);
        }
        return result;
    }
    //遍历每一行获取具体单元格中的数据并写入Bean对象中返回
    protected User getBeanFromRow(Row row){
        if(row == null)
            return null;
        int index = row.getFirstCellNum();
        Cell cell = row.getCell(index);
        if(cell != null){
            User user = new User();
            user.setUsername(cell.getStringCellValue());
            index++;

            cell = row.getCell(index);
            user.setPassword(cell.getStringCellValue());

            return user;
        }
        return null;
    }

}
