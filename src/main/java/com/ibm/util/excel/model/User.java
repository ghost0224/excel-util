package com.ibm.util.excel.model;

import lombok.Data;

import java.util.ArrayList;
import java.util.List;

/**
 * @author yiqing
 * @description 用户信息
 * @date 29/11/30
 */
@Data
public class User {

    private String name;
    private String company;
    private List<Punch> punchList = new ArrayList<>();
    private String floor;

}
