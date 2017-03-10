/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.victorlaerte.myexcelparser.util;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.util.IOUtils;
import org.json.JSONException;
import org.json.JSONObject;
import org.json.JSONTokener;

/**
 *
 * @author victor 
 */
public class MyExcelParserUtil {
    
    public static synchronized Map<String, String> getVersionMap() throws JSONException {

        Map<String, String> version = new HashMap<String, String>();

        JSONTokener jsonTokener = new JSONTokener(getJsonVersion());
        JSONObject root = new JSONObject(jsonTokener);
        
        JSONObject jsonVersion = (JSONObject) root.get("version");
        String userCode = jsonVersion.getString("userCode");
        String code = jsonVersion.getString("code");

        version.put("userCode", userCode);
        version.put("code", code);

        if (jsonVersion.has("message")) {

            String message = jsonVersion.getString("message");
            
            if (!message.trim().equals("")){
                version.put("message", message);
            }
        }

        return version;
    }

    private static synchronized String getJsonVersion() {

        StringBuilder json = new StringBuilder();
        int data;
        char c;

        InputStream file = MyExcelParserUtil.class.getResourceAsStream("version.json");

        try {
            while ((data = file.read()) != -1) {

                c = (char) data;

                json.append(c);
            }

        } catch (Exception e) {
            
            e.printStackTrace();
        } finally {

            IOUtils.closeQuietly(file);
        }
        return json.toString();
    }
}
