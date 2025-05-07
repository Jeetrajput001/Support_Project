package com.decimal.support.controller;

import com.decimal.support.service.DailyReport;
import com.decimal.support.service.WeeklyReport;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;

@RestController
public class Controller {

    @Autowired
    private WeeklyReport weeklyReport;
    @Autowired
    private DailyReport dailyReport;

    @GetMapping("/weekly-report")
    public String monthlyReport() {
        String issueTypeToFilter = "Service Request"; // Specify the issue type you want to filter
        try {
            weeklyReport.filterIssuesAndCollectIDs(issueTypeToFilter);
            return "Filtered IDs have been written to the new Excel file.";
        } catch (IOException e) {
            e.printStackTrace();
            return "Error occurred while filtering issues.";
        }
    }

        @GetMapping("/daily-report")
        public String filterIssues() {
            String issueTypeToFilter = "Service Request"; // Specify the issue type you want to filter
            try {
                dailyReport.filterIssuesAndCollectIDs(issueTypeToFilter);
                return "Filtered IDs have been written to the new Excel file.";
            } catch (IOException e) {
                e.printStackTrace();
                return "Error occurred while filtering issues.";
            }
        }
    }
