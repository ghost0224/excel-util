package com.ibm.util.excel.api;

import com.ibm.util.excel.model.User;

import java.io.IOException;
import java.util.List;

public interface DataAPI {

    List<User> getData(String path) throws IOException;

}
