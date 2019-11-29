package com.ibm.util.excel.resolver;

import com.ibm.util.excel.api.DataAPI;
import com.ibm.util.excel.common.Constant;
import com.ibm.util.excel.model.Punch;
import com.ibm.util.excel.model.User;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author yiqing
 * @description 样式二解析器
 * @date 29/11/30
 */
public class Style2Data implements DataAPI {

    @Override
    public List<User> getData(String path) throws IOException {
        InputStream input = new FileInputStream(path);
        Workbook wb;
        wb = new XSSFWorkbook(input);
        //样式二
        int sheetId = 1;
        Sheet sheet = wb.getSheetAt(sheetId);
        List<User> userList = new ArrayList<>();
        List<Punch> punchList = new ArrayList<>();
        //获取日期前缀
        Row rowDate = sheet.getRow(2);
        String date = rowDate.getCell(25).getStringCellValue();
        int idx = date.indexOf("～") + 1;
        date = date.substring(idx, idx + 8);
        int lastRowNum = sheet.getLastRowNum();
        User user = null;
        for (int i = 4; i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            Cell cellTitle = row.getCell(10);
            String title;
            try {
                 if(null != user && null != punchList && (cellTitle.getCellType() == XSSFCell.CELL_TYPE_STRING
                         || cellTitle.getCellType() == XSSFCell.CELL_TYPE_BLANK)
                         && cellTitle.getStringCellValue().indexOf("姓名") <0 ) {
                    //获取打卡信息
                     int t=1;
                     for (int j = 0; j < punchList.size(); j++) {
                        Cell cellDay = row.getCell(t);
                        if (null != cellDay && null != cellDay.getStringCellValue() && cellDay.getStringCellValue().length() > 0) {
                            Punch punch = punchList.get(j);
                            String time = cellDay.getStringCellValue();
                            time = time.replaceAll("\n", "");
                            time = time.trim();
                            punch.setData(time);
                            String start = "";
                            String end = "";
                            if (time.length() >= 0) {
                                start = time.substring(0, 5);
                            }
                            if (time.length() >= 6) {
                                end = time.substring(5);
                            }
                            if (null == punch.getStartTime()) {
                                punch.setStartTime(start);
                            } else {
                                //解决一个人一天里多行打卡时间的问题
                                if (start.length() > 0) {
                                    punch.setEndTime(start);
                                }
                            }
                            if (end.length() > 0) {
                                punch.setEndTime(end);
                            }
                            punch.setDiff();
                        }
                        t++;
                     }
                } else if(cellTitle.getCellType() == XSSFCell.CELL_TYPE_STRING
                        && cellTitle.getStringCellValue().indexOf("姓名") >=0 ) {
                     if (null != user) {
                         user.setPunchList(punchList);
                         userList.add(user);
                     }
                     //获取用户信息
                    title = cellTitle.getStringCellValue();
                    Cell cellName = row.getCell(11);
                    String name = cellName.getStringCellValue();
                    Cell cellCompany = row.getCell(18);
                    String company = cellCompany.getStringCellValue();
                    user = new User();
                    user.setCompany(company.toUpperCase());
                    user.setName(name.replaceAll("20",""));
                    user.setFloor(Constant.FLOOR_LIST[1]);
                    punchList = new ArrayList<>();
                } else if(null != user && cellTitle.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
                    //获取打卡日期
                    for (int j = 1; j < row.getLastCellNum(); j++) {
                        Cell cellDay = row.getCell(j);
                        if (null != cellDay) {
                            Integer day = ((Double)cellDay.getNumericCellValue()).intValue();
                            if (day > 0) {
                                Punch punch = new Punch();
                                punch.setDate(date, day);
                                punchList.add(punch);
                            }
                        }
                    }
                }
            } catch (RuntimeException ex) {
                ex.printStackTrace();
            }
        }
        return userList;
    }

}
