/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package com.victorlaerte.myexcelparser.service;

import com.victorlaerte.myexcelparser.FXMLDocumentController;
import com.victorlaerte.myexcelparser.util.HttpUtil;
import com.victorlaerte.myexcelparser.util.MyExcelParserUtil;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.json.JSONObject;

/**
 *
 * @author facilit-denv-01
 */
public class GetVersionTask extends AsyncTask {

    private boolean result = false;
    private Map<String, String> version;
    private FXMLDocumentController controller;
    private String message;

    public GetVersionTask(FXMLDocumentController controller) {
        this.controller = controller;
    }

    @Override
    void onPreExecute() {
    }

    @Override
    void doInBackground() {

        String url = "https://dl.dropboxusercontent.com/u/24099655/version-excel-app.json";

        try {
            JSONObject json = HttpUtil.getJSON(false, url, null, null, null);

            if (json.getInt("statusCode") == 200) {

                JSONObject body = json.getJSONObject("body");
                JSONObject jsonVersion = body.getJSONObject("version");
                String code = jsonVersion.getString("code");

                if (jsonVersion.has("message")) {
                    
                    message = jsonVersion.getString("message");
                }

                version = MyExcelParserUtil.getVersionMap();

                if (version.get("code") != null && code != null) {

                    int codeToCompare = Integer.parseInt(code);
                    int currentCode = Integer.parseInt(version.get("code"));

                    if (codeToCompare <= currentCode) {

                        System.out.println("Nenhuma nova versão encontrada");
                        result = false;

                    } else if (codeToCompare > currentCode) {

                        System.out.println("Nova versão encontrada");
                        result = true;
                    }
                }
            }

        } catch (Exception e) {
            Logger.getLogger(GetVersionTask.class.getName()).log(Level.SEVERE, e.getMessage());
        }
    }

    @Override
    void onPostExecute() {

        if (result) {

            controller.newVersionFound(version, message);
        }
    }
}
