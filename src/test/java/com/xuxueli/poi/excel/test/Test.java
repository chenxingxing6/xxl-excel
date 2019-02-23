package com.xuxueli.poi.excel.test;

import com.alibaba.fastjson.JSON;
import com.xuxueli.poi.excel.ExcelExportUtil;
import com.xuxueli.poi.excel.ExcelImportUtil;
import com.xuxueli.poi.excel.test.model.ShopDTO;
import com.xuxueli.poi.excel.test.model.UserDTO;
import org.apache.poi.ss.usermodel.Row;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * FUN Test
 *
 * @author xuxueli 2017-09-08 22:41:19
 */
public class Test {

    public static void main(String[] args) {

        /**
         * Mock数据，Java对象列表
         */
        List<ShopDTO> shopDTOList = new ArrayList<ShopDTO>();
        List<UserDTO> userDTOList = new ArrayList<>();
        for (int i = 0; i < 100; i++) {
            ShopDTO shop = new ShopDTO(true, "商户"+i, (short) i, 1000+i, 10000+i, (float) (1000+i), (double) (10000+i), new Date());
            shopDTOList.add(shop);
        }
        for (int i = 0; i < 100; i++) {
            UserDTO userDTO = new UserDTO();
            userDTO.setUsername("用户"+ i);
            userDTO.setAge(i);
            userDTOList.add(userDTO);
        }
        String filePath = "/Users/cxx/Downloads/demo-sheet.xls";

        /**
         * Excel导出：Object 转换为 Excel
         */
        ExcelExportUtil util = new ExcelExportUtil();
        Row cells = util.addRow(0);
        util.addCell(cells, 1, "title");
        util.exportToFile(filePath, shopDTOList, userDTOList);

        /**
         * Excel导入：Excel 转换为 Object
          */
      /*  List<Object> list = ExcelImportUtil.importExcel(filePath, ShopDTO.class);
        System.out.println(JSON.toJSONString(list));

        List<Object> list1 = ExcelImportUtil.importExcel(filePath, UserDTO.class);
        System.out.println(JSON.toJSONString(list1));*/

    }

}
