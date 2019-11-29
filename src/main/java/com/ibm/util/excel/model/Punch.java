package com.ibm.util.excel.model;

import lombok.Data;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @author yiqing
 * @description 打卡记录
 * @date 29/11/30
 */
@Data
public class Punch {

    private SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm");

    private String data;
    private String date;
    private String startTime;
    private String endTime;
    private String diff;

    public void setDate(String date, Integer day) {
        if (day < 10) {
            this.date = date + "0" + day;
        } else {
            this.date = date + day;
        }
    }

    public void setDiff() {
        if (null != startTime && null != endTime) {
            try {
                Date start = sdf.parse(date + " " + startTime);
                Date end = sdf.parse(date + " " + endTime);
                long diff = end.getTime() - start.getTime();
                long hour = diff / (1000 * 3600);
                long min = diff % (1000 * 3600);
                min = min / (1000 * 60);
                String hourStr;
                if (hour < 10) {
                    hourStr = "0" + hour;
                } else {
                    hourStr = hour + "";
                }
                String minStr;
                if (min < 10) {
                    minStr = "0" + min;
                } else {
                    minStr = min + "";
                }
                this.diff = hourStr + ":" + minStr;
            } catch (ParseException e) {
                e.printStackTrace();
            }
        }
    }

}
