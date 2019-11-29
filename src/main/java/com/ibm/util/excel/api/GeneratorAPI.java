package com.ibm.util.excel.api;

import com.ibm.util.excel.model.User;

import java.util.List;

public interface GeneratorAPI {

    void buildView(String path, List<User> userRecords);

}
