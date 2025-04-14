package com.decimal.support.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

@Service
public class ExcelService {

    public void createNewExcelWithSelectedColumns(String inputFilePath, String outputFilePath) throws IOException {
        try (FileInputStream inputStream = new FileInputStream(inputFilePath);
             Workbook workbook = new XSSFWorkbook(inputStream);
             Workbook newWorkbook = new XSSFWorkbook()) {

            Sheet originalSheet = workbook.getSheetAt(0);
            Sheet newSheet = newWorkbook.createSheet("Filtered Data");

            for (Row row : originalSheet) {
                Row newRow = newSheet.createRow(row.getRowNum());
                // Assuming you want to copy columns 0, 1, and 2 (first three columns)
                copyCellValue(row.getCell(0), newRow.createCell(0));
                copyCellValue(row.getCell(1), newRow.createCell(1));
                copyCellValue(row.getCell(2), newRow.createCell(2));
            }

            try (FileOutputStream outputStream = new FileOutputStream(outputFilePath)) {
                newWorkbook.write(outputStream);
            }
        }
    }
    private void copyCellValue(Cell sourceCell, Cell targetCell) {
        if (sourceCell != null) {
            switch (sourceCell.getCellType()) {
                case STRING:
                    targetCell.setCellValue(sourceCell.getStringCellValue());
                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(sourceCell)) {
                        targetCell.setCellValue(sourceCell.getDateCellValue());
                    } else {
                        targetCell.setCellValue(sourceCell.getNumericCellValue());
                    }
                    break;
                case BOOLEAN:
                    targetCell.setCellValue(sourceCell.getBooleanCellValue());
                    break;
                case FORMULA:
                    targetCell.setCellFormula(sourceCell.getCellFormula());
                    break;
                case BLANK:
                    targetCell.setBlank();
                    break;
                default:
                    targetCell.setCellValue(sourceCell.toString());
                    break;
            }
        }
    }

}