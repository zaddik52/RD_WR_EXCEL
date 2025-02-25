package com.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelProcessor {

    public static void main(String[] args) {
        String inputFilePath = "list_all.xlsx";
        String outputFilePath = "processed_list_all.xlsx";

        try {
            FileInputStream fis = new FileInputStream(inputFilePath);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet inputSheet = workbook.getSheetAt(0); // קריאת ה-SHEET הראשון

            // יצירת SHEET חדש לעיבוד הנתונים
            Sheet outputSheet = workbook.createSheet("ProcessedData");

            // קריאת הנתונים מה-SHEET הראשון ועיבודם
            for (Row row : inputSheet) {
                Row newRow = outputSheet.createRow(row.getRowNum());
                for (Cell cell : row) {
                    Cell newCell = newRow.createCell(cell.getColumnIndex());
                    // עיבוד הנתונים (לדוגמה: הכפלת ערכים במספר מסוים)
                    if (cell.getCellType() == CellType.NUMERIC) {
                        newCell.setCellValue(cell.getNumericCellValue() * 2);
                    } else if (cell.getCellType() == CellType.STRING) {
                        newCell.setCellValue(cell.getStringCellValue().toUpperCase());
                    } else {
                        newCell.setCellValue(cell.toString());
                    }
                }
            }

            // שמירת הקובץ לאחר העיבוד
            FileOutputStream fos = new FileOutputStream(outputFilePath);
            workbook.write(fos);
            fos.close();
            workbook.close();
            fis.close();

            System.out.println("Data processed and written to " + outputFilePath);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}