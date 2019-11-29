package com.ibm.util.excel;

import com.ibm.util.excel.api.DataAPI;
import com.ibm.util.excel.api.GeneratorAPI;
import com.ibm.util.excel.generator.DefaultGenerator;
import com.ibm.util.excel.model.Punch;
import com.ibm.util.excel.model.User;
import com.ibm.util.excel.resolver.Style1Data;
import com.ibm.util.excel.resolver.Style2Data;

import java.io.IOException;
import java.util.List;

/**
 * @author yiqing
 * @description Starter
 * @date 29/11/30
 */
public class Starter {

    public static void main(String[] args) throws IOException {
        String input = args[0];
        String output = args[1];
        DataAPI dataAPI1 = new Style1Data();
        List<User> userRecords = dataAPI1.getData(input);
        DataAPI dataAPI2 = new Style2Data();
        List<User> userRecords2 = dataAPI2.getData(input);
        userRecords.addAll(userRecords2);
        for (User m : userRecords) {
            System.out.println(m.getName() + ":" + m.getCompany() + "," + m.getFloor());
            for (Punch p : m.getPunchList()) {
                System.out.println(p.getDate() + "-" + p.getStartTime() + "," + p.getEndTime() + "-" + p.getDiff());
            }
        }
        GeneratorAPI generator = new DefaultGenerator();
        generator.buildView(output, userRecords);
    }



}
