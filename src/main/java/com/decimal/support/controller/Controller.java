package com.decimal.support.controller;

import com.decimal.support.service.ExcelService;
import com.decimal.support.service.ExcelServices;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;

@RestController
public class Controller {

    @Autowired
    private ExcelService excelService;
    @Autowired
    private ExcelServices excelServices;

    @GetMapping("/create-excel")
    public String createExcel() {
        String inputFilePath = "/home/decimal/Documents/12-03-2025.xlsx"; // Update with your input file path
        String outputFilePath = "/home/decimal/Documents/spring.xlsx"; // Update with your desired output file path

        try {
            excelService.createNewExcelWithSelectedColumns(inputFilePath, outputFilePath);
            return "New Excel file created successfully!";
        } catch (IOException e) {
            e.printStackTrace();
            return "Error occurred while creating the Excel file.";
        }
    }

        @GetMapping("/filter-issues")
        public String filterIssues() {
            String issueTypeToFilter = "Service Request"; // Specify the issue type you want to filter
            try {
                excelServices.filterIssuesAndCollectIDs(issueTypeToFilter);
                return "Filtered IDs have been written to the new Excel file.";
            } catch (IOException e) {
                e.printStackTrace();
                return "Error occurred while filtering issues.";
            }
        }
    }
