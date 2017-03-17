package com.victorlaerte.myexcelparser;

import com.victorlaerte.myexcelparser.util.Dialog;
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.ResourceBundle;
import javafx.application.Platform;
import javafx.event.ActionEvent;
import javafx.event.Event;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.stage.DirectoryChooser;
import javafx.stage.Stage;
import javax.swing.filechooser.FileSystemView;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
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
    
    private final String DEFAULT_DIR = FileSystemView.getFileSystemView().getDefaultDirectory().getPath();
    private final String DEFAULT_IDENTIFIER = "BSE";
    private final String EXTENSION = ".XLS";
    private final String DEFAULT_SHEET_NAME = "BMP";
    private final String ITEM = "ITEM";
    private final String QTY = "QUANTIDADE";
    private final String PREFIX_OUTPUT_FILE_NAME = "Total - ";
    private final int MAX_ATTEMPT_GET_CELL = 4;
    private Stage stage;
    
    @FXML
    private TextField dirTxtField;
    @FXML
    private TextField identifierTxtField;
    @FXML
    private TextField sheetNameTxtField;
    
    @FXML
    private void handleButtonStart(ActionEvent event) {
        
        File dir = new File(dirTxtField.getText());
        
        List<File> xlsFileList = getExcelFiles(dir);
        
        Map<String, Integer> tableKeyValue = getTableKeyValue(xlsFileList);
        
        createNewExcelFile(dir, tableKeyValue);
    }
    
    @FXML
    private void handleButtonChooseDir(ActionEvent event) {
        
        DirectoryChooser chooser = new DirectoryChooser();
        chooser.setTitle("Selecione um diretório para iniciar");
        File selectedDirectory = chooser.showDialog(stage);
        
        if (selectedDirectory != null) {
            
            dirTxtField.setText(selectedDirectory.getAbsolutePath());
        }
    }
    
    public void setStage(Stage stage) {
        this.stage = stage;
    }
    
    @Override
    public void initialize(URL url, ResourceBundle rb) {
        
        dirTxtField.setText(DEFAULT_DIR);
        identifierTxtField.setText(DEFAULT_IDENTIFIER);
        sheetNameTxtField.setText(DEFAULT_SHEET_NAME);
        
    }

    private void createNewExcelFile(File dir, Map<String, Integer> tableKeyValue) {
        
        HSSFWorkbook workbook = new HSSFWorkbook();
        
        HSSFSheet sheet = workbook.createSheet(sheetNameTxtField.getText());
        
        int rowNum = 0;
        Row row = sheet.createRow(rowNum++);
        
        createHeader(row);
        
        for (Map.Entry<String, Integer> entry : tableKeyValue.entrySet()) {
            
            HSSFRow currentRow = sheet.createRow(rowNum++);
            
            String key = entry.getKey();
            
            Integer value = entry.getValue();
            
            System.out.println(key + " " + value);
            
            HSSFCell firstColumn = currentRow.createCell(0);
            firstColumn.setCellValue(key);
            
            HSSFCell secondColumn = currentRow.createCell(1);
            secondColumn.setCellValue(value);
        }
        
        try {
            writeFile(dir, workbook);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void writeFile(File dir, HSSFWorkbook workbook) throws IOException {
            
        FileOutputStream outputStream = null;
                
        try {
            
            Date now = new Date();
            SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
            String formatedDate = sdf.format(now);
            
            String outputFileName = dir + File.separator + PREFIX_OUTPUT_FILE_NAME + formatedDate;
            
            outputFileName = getUniqueFileName(outputFileName, 1);
            
            outputStream = new FileOutputStream(outputFileName);
            workbook.write(outputStream);
            
        } finally {
            
            if (outputStream != null) {
                
                outputStream.close();
            }
            workbook.close();
        }
    }
    
    private String getUniqueFileName(String originalFilePath, int equalFileCount) {
     
        String newFilePath = originalFilePath + EXTENSION.toLowerCase();
        
        File file = new File(newFilePath);
        
        if (file.exists()) {
            
            newFilePath = originalFilePath + "(" + equalFileCount + ")" + EXTENSION.toLowerCase();
            
            File newFile = new File(newFilePath);
            
            if (newFile.exists()) {
                
                newFilePath = getUniqueFileName(originalFilePath, ++equalFileCount);
            }
        }
        
        return newFilePath;
    }

    private void createHeader(Row row) {
        
        Cell firstColumn = row.createCell(0);
        firstColumn.setCellValue(ITEM);
        
        Cell secondColumn = row.createCell(1);
        secondColumn.setCellValue(QTY);
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
                
                Sheet sheet = workbook.getSheet(sheetNameTxtField.getText());
                
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
                
                if (fileName.contains(identifierTxtField.getText()) && fileName.contains(EXTENSION)) {
                    
                    xlsFileList.add(listFile);
                }
            }
        }
        
        return xlsFileList;
    }
    
       public void newVersionFound(final Map<String, String> version, String message) {

        String updateMessage = "";

        if (message != null && !message.trim().equals("")) {

            updateMessage = message;
        }

           Dialog.buildConfirmation("Nova Versão", "Existe uma nova versão disponível. Deseja instala-la agora? " + updateMessage)
                .addYesButton(new EventHandler() {
                    
                    @Override
                    public void handle(Event t) {
                        
                        String url = "https://dl.dropboxusercontent.com/content_link/QHP2de32UAKpNA01CvbLeQkLxNYAUipd6iyfsRP3sIVScd47aWEmU7jpYRH1fb80/file?dl=1";

                        if (!url.equals("")) {

                            URI uri;
                            try {
                                uri = new URI(url);
                                Desktop.getDesktop().browse(uri);
                                Platform.exit();
                            } catch (URISyntaxException | IOException ex) {
                                ex.printStackTrace();
                            }
                        }
                    }
                })
                .addNoButton(null)
                .build()
                .show();
    }
}
