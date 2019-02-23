package com.xuxueli.poi.excel.test.model;

import com.xuxueli.poi.excel.annotation.ExcelField;
import com.xuxueli.poi.excel.annotation.ExcelSheet;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.util.Date;

/**
 * Java Object To Excel
 *
 * @author xuxueli 2017-09-12 11:20:02
 */
@ExcelSheet(name = "用户列表", headColor = HSSFColor.HSSFColorPredefined.BRIGHT_GREEN)
public class UserDTO {

    @ExcelField(name = "用户名", align = HorizontalAlignment.CENTER)
    private String username;

    @ExcelField(name = "年龄")
    private int age;

    public UserDTO() {
    }

    public String getUsername() {
        return username;
    }

    public void setUsername(String username) {
        this.username = username;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }
}
