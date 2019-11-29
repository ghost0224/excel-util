package com.ibm.util.excel.resolver;

import com.ibm.util.excel.api.DataAPI;
import com.ibm.util.excel.common.Constant;
import com.ibm.util.excel.model.Punch;
import com.ibm.util.excel.model.User;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author yiqing
 * @description 样式一解析器
 * @date 29/11/30
 */
public class Style1Data implements DataAPI {

    @Override
    public List<User> getData(String path) throws IOException {
        InputStream input = new FileInputStream(path);
        Workbook wb;
        wb = new XSSFWorkbook(input);
        //样式一
        int sheetId = 0;
        Sheet sheet = wb.getSheetAt(sheetId);
        List<Punch> punchList = new ArrayList<>();
        //获取日期前缀
        Row rowDate = sheet.getRow(2);
        String date = rowDate.getCell(2).getStringCellValue();
        date = date.substring(0, 8);
        //获取日期后缀
        Row rowDay = sheet.getRow(3);
        for (int d = 0; d < rowDay.getLastCellNum(); d++) {
            Cell cellDay = rowDay.getCell(d);
            Integer day = ((Double)cellDay.getNumericCellValue()).intValue();
            if (day > 0) {
                Punch punch = new Punch();
                punch.setDate(date, day);
                punchList.add(punch);
            }
        }
        List<User> userList = new ArrayList<>();
        int lastRowNum = sheet.getLastRowNum();
        User user = null;
        for (int i = 4; i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            //获取用户信息
            if (i % 2 == 0) {
                Cell cellName = row.getCell(10);
                String name = cellName.getStringCellValue();
                if (null != name && name.length() > 0) {
                    user = new User();
                    user.setName(name);
                    user.setFloor(Constant.FLOOR_LIST[0]);
                }
                Cell cellCompany = row.getCell(20);
                String company = cellCompany.getStringCellValue();
                if (null != company && company.length() > 0 && null != user) {
                    user.setCompany(company.toUpperCase());
                }
            } else {
                //获取打卡信息
                List<Punch> punches = new ArrayList<>();
                for (int j = 0; j < punchList.size(); j++) {
                    Punch punch = new Punch();
                    Cell cell = row.getCell(j);
                    String time = cell.getStringCellValue();
                    if (null != time && time.trim().length() > 0) {
                        time = time.trim();
                        punch.setData(time);
                        punch.setDate(date, j+1);
                        String start = time.substring(0, 5);
                        punch.setStartTime(start);
                        punch.setDiff("N/A");
                        if (time.length() > 5) {
                            int endIdx = time.length();
                            int startIdx = endIdx - 5;
                            String end = time.substring(startIdx, endIdx);
                            punch.setEndTime(end);
                            punch.setDiff();
                        }
                    }
                    punches.add(punch);
                }
                if (null != user) {
                    user.setPunchList(punches);
                    userList.add(user);
                    user = null;
                }
            }
        }
        return userList;
    }
}
