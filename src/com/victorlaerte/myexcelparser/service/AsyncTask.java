/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package com.victorlaerte.myexcelparser.service;

import javafx.application.Platform;

/**
 *
 * @author facilit-denv-01
 */
public abstract class AsyncTask {

    abstract void onPreExecute();

    abstract void doInBackground();

    abstract void onPostExecute();

    public void execute() {

        Platform.runLater(new Runnable() {
            @Override
            public void run() {
                onPreExecute();

                new Thread(new Runnable() {
                    @Override
                    public void run() {

                        doInBackground();

                        Platform.runLater(new Runnable() {
                            @Override
                            public void run() {
                                onPostExecute();
                            }
                        });
                    }
                }).start();
            }
        });
    }
}