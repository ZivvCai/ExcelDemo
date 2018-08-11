package com.czw.testDemo;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class TestWriteExcel {

    @Test
    public void testWrite() throws InvalidFormatException, IOException, ClassNotFoundException {
        File file = new File("F:/testWrite.xlsx");
        if(file.exists())
            file.delete();
        WriteExcel writeExcel = new WriteExcel(file);

        User user1 = new User("aaa","123");
        User user2 = new User("bbb","456");
        User user3 = new User("ccc","789");

        List<User> users = new ArrayList<User>();
        users.add(user1);
        users.add(user2);
        users.add(user3);

        writeExcel.write(users);
        writeExcel.execute();

        if(file.exists()){
            System.out.println("存在");
        }

    }

}
