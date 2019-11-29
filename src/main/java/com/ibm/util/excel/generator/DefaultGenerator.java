package com.ibm.util.excel.generator;

import com.ibm.util.excel.api.GeneratorAPI;
import com.ibm.util.excel.model.Punch;
import com.ibm.util.excel.model.User;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

/**
 * @author yiqing
 * @description 默认生成器
 * @date 29/11/30
 */
public class DefaultGenerator implements GeneratorAPI {

    @Override
    public void buildView(String path, List<User> userRecords) {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("考勤统计");
        Row headRow = sheet.createRow(0);
        Cell headCell1 = headRow.createCell(0);
        headCell1.setCellValue("4层工号");
        Cell headCell2 = headRow.createCell(1);
        headCell2.setCellValue("4层姓名");
        Cell headCell3 = headRow.createCell(2);
        headCell3.setCellValue("20层工号");
        Cell headCell4 = headRow.createCell(3);
        headCell4.setCellValue("20层姓名");
        Cell headCell5 = headRow.createCell(4);
        headCell5.setCellValue("公司");
        Cell headCell6 = headRow.createCell(5);
        headCell6.setCellValue("分组");
        Cell headCell7 = headRow.createCell(6);
        headCell7.setCellValue("考勤日期");
        Cell headCell8 = headRow.createCell(7);
        headCell8.setCellValue("考勤记录");
        Cell headCell9 = headRow.createCell(8);
        headCell9.setCellValue("考勤开始");
        Cell headCell10 = headRow.createCell(9);
        headCell10.setCellValue("考勤结束");
        Cell headCell11 = headRow.createCell(10);
        headCell11.setCellValue("考勤时间");

        int i = 1;
        for (User user : userRecords) {
            for (Punch punch : user.getPunchList()) {
                Row row = sheet.createRow(i++);
                Cell cell1 = row.createCell(0);
                Cell cell2 = row.createCell(1);
                if(user.getFloor().equalsIgnoreCase("4层")) {
                    cell2.setCellValue(user.getName());
                    if (null != user.getName()) {
                        sheet.setColumnWidth(1, user.getName().getBytes().length*2*256);
                    }
                }
                Cell cell3 = row.createCell(2);
                Cell cell4 = row.createCell(3);
                if(user.getFloor().equalsIgnoreCase("20层")) {
                    cell4.setCellValue(user.getName());
                    if (null != user.getName()) {
                        sheet.setColumnWidth(3, user.getName().getBytes().length*2*256);
                    }
                }
                Cell cell5 = row.createCell(4);
                cell5.setCellValue(user.getCompany());
                if (null != user.getCompany()) {
                    sheet.setColumnWidth(4, user.getCompany().getBytes().length*2*256);
                }
                Cell cell6 = row.createCell(5);
                Cell cell7 = row.createCell(6);
                cell7.setCellValue(punch.getDate());
                if (null != punch.getDate()) {
                    sheet.setColumnWidth(6, punch.getDate().getBytes().length*2*256);
                }
                Cell cell8 = row.createCell(7);
                cell8.setCellValue(punch.getData());
                if (null != punch.getData()) {
                    sheet.setColumnWidth(7, punch.getData().getBytes().length*2*256);
                }
                Cell cell9 = row.createCell(8);
                cell9.setCellValue(punch.getStartTime());
                if (null != punch.getStartTime()) {
                    sheet.setColumnWidth(8, punch.getStartTime().getBytes().length*2*256);
                }
                Cell cell10 = row.createCell(9);
                cell10.setCellValue(punch.getEndTime());
                if(null != punch.getEndTime()) {
                    sheet.setColumnWidth(9, punch.getEndTime().getBytes().length*2*256);
                }
                Cell cell11 = row.createCell(10);
                cell11.setCellValue(punch.getDiff());
                if (null != punch.getDiff()) {
                    sheet.setColumnWidth(10, punch.getDiff().getBytes().length*2*256);
                }
            }
        }

        try  (OutputStream fileOut = new FileOutputStream(path)) {
            wb.write(fileOut);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
