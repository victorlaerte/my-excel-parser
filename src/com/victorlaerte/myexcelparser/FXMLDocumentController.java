package com.victorlaerte.myexcelparser;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.ResourceBundle;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Label;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 *
 * @author Victor Oliveira
 */
public class FXMLDocumentController implements Initializable {
    
    private String directory = "C:\\Users\\victo\\Desktop\\bse";
    private final String IDENTIFIER = "BSE";
    private final String EXTENSION = ".XLS";
    private final String SHEET_NAME = "BMP";
    private final String ITEM = "ITEM";
    private final int MAX_ATTEMPT_GET_CELL = 4;
    
    @FXML
    private Label label;
    
    @FXML
    private void handleButtonAction(ActionEvent event) {
        
        
    }
    
    @Override
    public void initialize(URL url, ResourceBundle rb) {
        
        File dir = new File(directory);
        
        List<File> xlsFileList = getExcelFiles(dir);
        
        Map<String, Integer> tableKeyValue = getTableKeyValue(xlsFileList);
        
        for (Map.Entry<String, Integer> entry : tableKeyValue.entrySet()) {
            
            String key = entry.getKey();
            
            Integer value = entry.getValue();
            
            System.out.println(key + " " + value);
        }
    }

    private Map<String, Integer> getTableKeyValue(List<File> xlsFileList) throws NumberFormatException, EncryptedDocumentException {
        
        Map<String, Integer> tableKeyValue = new LinkedHashMap<String, Integer>();
                
        boolean startSum = false;
        
        for (File xlsFile : xlsFileList) {
        
            try {
                
                System.out.println("###########################################");
                System.out.println("File " + xlsFile.getName());
                
                FileInputStream fsXls = new FileInputStream(xlsFile);
                
                Workbook workbook = WorkbookFactory.create(fsXls);
                
                Sheet sheet = workbook.getSheet(SHEET_NAME);
                
                Iterator<Row> iterator = sheet.iterator();
                
                while (iterator.hasNext()) {
                    
                    Row currentRow = iterator.next();
                    
                    Cell firstCell = currentRow.getCell(0);
                    
                    int attempts = 1;
                    
                    if (firstCell != null && firstCell.getCellTypeEnum() == CellType.STRING) {
                        
                        String stringCellValue = firstCell.getStringCellValue();

                        if (startSum == false && stringCellValue.equalsIgnoreCase(ITEM)) {

                            startSum = true;
                            
                        } else if (startSum) {
                            
                            Cell secondCell = currentRow.getCell(attempts);
                            
                            if (secondCell != null) {
                                
                                int cellValue = getCellValue(secondCell, attempts, currentRow);
                            
                                addToTableKeyValueMap(tableKeyValue, stringCellValue, cellValue);
                            }
                        }
                    }
                }
                
            } catch (Exception e) {
                e.printStackTrace();
            } finally  {
                startSum = false;
                System.out.println("###########################################");
            }
        }
        
        return tableKeyValue;
    }

    private int getCellValue(Cell secondCell, int attempts, Row currentRow) throws NumberFormatException {
        
        int secondCellValue = 0;
        CellType cellTypeEnum = secondCell.getCellTypeEnum();
        
        if (cellTypeEnum == CellType.NUMERIC) {
            
            secondCellValue = (int) secondCell.getNumericCellValue();
            
        } else if (cellTypeEnum == CellType.STRING) {
            
            String secondCellValueStr = secondCell.toString();
            
            if (secondCellValueStr!=null && !secondCellValueStr.isEmpty()) {
                
                secondCellValue = Integer.valueOf(secondCellValueStr);
            }
            
        } else {
            
            while (cellTypeEnum != CellType.NUMERIC && attempts <= MAX_ATTEMPT_GET_CELL) {
                
                secondCell = currentRow.getCell(attempts);
                
                if (secondCell != null) {
                    
                    cellTypeEnum = secondCell.getCellTypeEnum();
                }
                
                attempts += 1;
            }
            
            if (cellTypeEnum == CellType.NUMERIC) {
                
                secondCellValue = (int) secondCell.getNumericCellValue();
            }
        }
        
        return secondCellValue;
    }

    private void addToTableKeyValueMap(Map<String, Integer> tableKeyValue, String stringCellValue, int secondCellValue) {
        
        if (tableKeyValue.containsKey(stringCellValue)) {
            
            int mapValue = tableKeyValue.get(stringCellValue);
            
            mapValue += secondCellValue;
            
            tableKeyValue.put(stringCellValue, mapValue);
            
            System.out.println(stringCellValue + " - " + mapValue);
            
        } else {
            
            tableKeyValue.put(stringCellValue, secondCellValue);
            
            System.out.println(stringCellValue + " - " + secondCellValue);
        }
    }

    private List<File> getExcelFiles(File dir) {
        
        List<File> xlsFileList = new ArrayList<File>();
        
        if (dir.exists() && dir.isDirectory()) {
            
            File[] listFiles = dir.listFiles();
            
            for (File listFile : listFiles) {
                
                String fileName = listFile.getName().toUpperCase();
                
                if (fileName.contains(IDENTIFIER) && fileName.contains(EXTENSION)) {
                    
                    xlsFileList.add(listFile);
                }
            }
        }
        
        return xlsFileList;
    }
}
