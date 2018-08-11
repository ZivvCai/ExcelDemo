package com.czw.testDemo;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.util.List;

public class TestReadExcel {

    @Test
    public void testRead() throws InvalidFormatException, IOException {
        File file = new File("F:/testRead.xlsx");
        ReadExcel readExcel = new ReadExcel();
        List<User> users = readExcel.readXLSX(file);
        for(User user : users){
            System.out.println(user.toString());
        }
        System.out.println(users.size());
    }

}
