package com.decimal.support.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

@Service
public class WeeklyReport {


    private String inputFilePath = "/home/decimal/Downloads/L2_10d_Report1905.xlsx";  //input file path
    private String outputFilePath = "/home/decimal/Support/Auto_Generated_Report/WeeklyReport/19-05-2025Weekly.xlsx"; //output file path
    private String grid = "/home/decimal/Downloads/Daily Report updated Grid.xlsx"; //grid input path
    List<String[]> collectedIDs = new ArrayList<>(); //main list which collects all the data from sheet and then write into new sheet


    public void filterIssuesAndCollectIDs(String issueTypeToFilter) throws IOException {

        Map<String, String> ClientNameMap = loadClientData();
        Map<String, String> EnglisnameMap = loadEnglishNameData();
        Map<String, Map<String, String[]>> subtasks = preprocessSubtasks(inputFilePath);
        List<String> activeStatus= Arrays.asList("Released","New","On Hold","Accepted in Roadmap","Testing In Progress","Awaiting Client Response","In Progress","Pending Release","TODO","Under Analysis");

        List<String> yellow=Arrays.asList("CSR","Infra","Developer","PM","Status CSR", "Assignee CSR","Created Date CSR", "Updated Date CSR",
                "Status Infra", "Assignee Infra", "Created Date Infra", "Updated Date Infra",
                "Status Developer", "Assignee Developer", "Created Date Developer", "Updated Date Developer",
                "Status PM", "Assignee PM","Created Date PM", "Updated Date PM");
        List<String> orange=Arrays.asList("DevOps","L2","BUG","Status DevOps", "Assignee DevOps", "Created Date DevOps", "Updated Date DevOps",
                "Status L2", "Assignee L2", "Created Date L2", "Updated Date L2",
                "Status Bug", "Assignee Bug", "Created Date Bug", "Updated Date Bug");
        List<String> srDetail=Arrays.asList("Issue Key", "Project Name", "Client", "English Name", "Reporters EmailId","Feasible Status ","Active PM/BUG", "Status", "Assignee",
                "Created Date", "Updated Date", "Components", "Labels","Priority","Tickets Aging");

        try (FileInputStream inputStream = new FileInputStream(inputFilePath);
             Workbook workbook = new XSSFWorkbook(inputStream);
             Workbook newWorkbook = new XSSFWorkbook()) {

            Sheet originalSheet = workbook.getSheetAt(0);
            Sheet newSheet = newWorkbook.createSheet("Weekly Report");

            // Font and Style for CSR (Yellow background)
            Font srFont = newWorkbook.createFont();
            srFont.setBold(true);
            srFont.setColor(IndexedColors.BLACK.getIndex());  // Black text color
            CellStyle srStyle = newWorkbook.createCellStyle();
            srStyle.setFont(srFont);
            srStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            srStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // Font and Style for CSR (Yellow background)
            Font csrFont = newWorkbook.createFont();
            csrFont.setBold(true);
            csrFont.setColor(IndexedColors.BLACK.getIndex());  // black text color
            CellStyle csrStyle = newWorkbook.createCellStyle();
            csrStyle.setFont(csrFont);
            csrStyle.setFillForegroundColor(IndexedColors.LIME.getIndex());
            csrStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // Font and Style for DevOps(Orange background)
            Font pmFont = newWorkbook.createFont();
            pmFont.setBold(true);
            pmFont.setColor(IndexedColors.BLACK.getIndex());  // Black text color
            CellStyle pmStyle = newWorkbook.createCellStyle();
            pmStyle.setFont(pmFont);
            pmStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
            pmStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // Create header row
            String[] headers = {
                    "Issue Key", "Project Name", "Client", "English Name", "Reporters EmailId","Feasible Status ","Active PM/BUG", "Status", "Assignee",
                    "Created Date", "Updated Date", "Components", "Labels","Priority", "Tickets Aging",
                    "CSR", "Status CSR", "Assignee CSR",
                    "Created Date CSR", "Updated Date CSR", "DevOps", "Status DevOps", "Assignee DevOps", "Created Date DevOps", "Updated Date DevOps",
                    "Infra", "Status Infra", "Assignee Infra", "Created Date Infra", "Updated Date Infra", "L2", "Status L2", "Assignee L2",
                    "Created Date L2", "Updated Date L2", "Developer", "Status Developer", "Assignee Developer", "Created Date Developer", "Updated Date Developer",
                    "BUG", "Status Bug", "Assignee Bug", "Created Date Bug", "Updated Date Bug", "PM", "Status PM", "Assignee PM",
                    "Created Date PM", "Updated Date PM"
            };

            Row headerRow = newSheet.createRow(0);

            for (int i = 0; i < headers.length; i++) {
                Cell headerCell = headerRow.createCell(i);
                headerCell.setCellValue(headers[i]);

                // Apply style based on header
                if (yellow.contains(headers[i])) {
                    headerCell.setCellStyle(csrStyle);
                } else if (orange.contains(headers[i])) {
                    headerCell.setCellStyle(pmStyle);
                }else if (srDetail.contains(headers[i])) {
                    headerCell.setCellStyle(srStyle);
                }
            }




            for (Row row : originalSheet) {
                Cell issueTypeCell = row.getCell(1); // Assuming issue type is in the first column
                Cell idCell = row.getCell(0);// Assuming ID is in the second column
                Cell projectnamecell = row.getCell(15);
                Cell reporterEmailcell = row.getCell(9);
                Cell createDatecell = row.getCell(5);
                Cell updateDatecell = row.getCell(6);
                Cell componentcell = row.getCell(13);
                Cell labelcell = row.getCell(14);
                Cell linkissued = row.getCell(12);
                Cell statusCell=row.getCell(4);
                Cell AsigneeCell=row.getCell(8);
                Cell prioritycell=row.getCell(2);


                if (issueTypeCell != null && idCell != null) {
                    String issueType = issueTypeCell.getStringCellValue();
                    if (issueType.equalsIgnoreCase(issueTypeToFilter)) {
                        String id = getCellValueAsString(idCell);
                        String projectName = getCellValueAsString(projectnamecell);
                        String AppName=getCellValueAsString(row.getCell(16));
                        String reporterEmail = getCellValueAsString(reporterEmailcell);
                        String status=getCellValueAsString(statusCell);
                        String asignee=getCellValueAsString(AsigneeCell);

                        String createddate = getCellValueAsString(createDatecell);
                        String updateDate = getCellValueAsString(updateDatecell);
                        String components = getCellValueAsString(componentcell);
                        String label = getCellValueAsString(labelcell);
                        String priority=getCellValueAsString(prioritycell);

                        String ticketAge="";
                        String clientcell = ClientNameMap.getOrDefault(AppName, "");
                        String englishcell = EnglisnameMap.getOrDefault(AppName, "");
                        String feasible =checkFeasibleStatus(status);
                        Map<String, String[]> taskMap = subtasks.getOrDefault(id, new HashMap<>());
                        String[] csrInfo = taskMap.getOrDefault("CSR", new String[5]);
                        String[] DevOpsInfo = taskMap.getOrDefault("DevOps", new String[5]);
                        String[] L2Info = taskMap.getOrDefault("L2-Debugging", new String[5]);
                        String[] infraInfo = taskMap.getOrDefault("Infra Operations", new String[5]);
                        String[] DeveloperInfo = taskMap.getOrDefault("Developer", new String[5]);
                        String[] BugInfo = taskMap.getOrDefault("BUG.", new String[5]);
                        String[] PMInfo = taskMap.getOrDefault("Product Manager", new String[5]);
                        String activePmBug="";
                        if (activeStatus.contains(PMInfo[1])) {
                            activePmBug = "PM";
                        }
                        else if (activeStatus.contains(BugInfo[1])){
                            activePmBug="BUG";
                        }
                        long days=0;
                        try{
                            DateTimeFormatter formatter= DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm");
                            LocalDate created =LocalDate.parse(createddate,formatter);
                            LocalDate updated =LocalDate.parse(updateDate,formatter);
                            LocalDate today =LocalDate.now();
                            if (feasible.equalsIgnoreCase("Active")||feasible.equalsIgnoreCase("ACR")){
                                for (LocalDate date = created; !date.isAfter(today); date = date.plusDays(1)) {
                                    DayOfWeek day = date.getDayOfWeek();
                                    if (day != DayOfWeek.SATURDAY && day != DayOfWeek.SUNDAY) {
                                        days++;
                                    }
                                }


                            } else if (feasible.equalsIgnoreCase("Rejected")||feasible.equalsIgnoreCase("Resolved")||feasible.equalsIgnoreCase("Closed")) {
                                for (LocalDate date = created; !date.isAfter(updated); date = date.plusDays(1)) {
                                    DayOfWeek day = date.getDayOfWeek();
                                    if (day != DayOfWeek.SATURDAY && day != DayOfWeek.SUNDAY) {
                                        days++;
                                    }
                                }


                            }

                        }catch (Exception e){
                            ticketAge="";

                        }
                        ticketAge=String.valueOf(days);

                        collectedIDs.add(new String[]{id, projectName, clientcell, englishcell,reporterEmail, feasible,activePmBug, status, asignee,createddate, updateDate, components, label,priority,ticketAge,
                                csrInfo[0],csrInfo[1],csrInfo[2],csrInfo[3],csrInfo[4],
                                DevOpsInfo[0],DevOpsInfo[1],DevOpsInfo[2],DevOpsInfo[3],DevOpsInfo[4],
                                infraInfo[0],infraInfo[1],infraInfo[2],infraInfo[3],infraInfo[4],
                                L2Info[0], L2Info[1], L2Info[2], L2Info[3], L2Info[4],
                                DeveloperInfo[0], DeveloperInfo[1], DeveloperInfo[2], DeveloperInfo[3], DeveloperInfo[4],
                                BugInfo[0],BugInfo[1],BugInfo[2],BugInfo[3],BugInfo[4],
                                PMInfo[0], PMInfo[1], PMInfo[2], PMInfo[3], PMInfo[4]


                        });

                    }


                }


            }

            // Write collected IDs to the new sheet
            int rowIndex = 1;
            for (String[] data : collectedIDs) {
                Row newRow = newSheet.createRow(rowIndex++);
                for (int colIndex = 0; colIndex < data.length; colIndex++) {
                    if (colIndex==6 && data[6].equals("")) {
                        continue;

                    }
                    if (colIndex==14) {
                        int age = Integer.parseInt(data[colIndex]);
                        newRow.createCell(colIndex).setCellValue(age);
                        continue;
                    }
                    newRow.createCell(colIndex).setCellValue(data[colIndex]);
                }
            }

            writeSummarySheet(newWorkbook);//pivot table creates
            try (FileOutputStream outputStream = new FileOutputStream(outputFilePath)) {
                newWorkbook.write(outputStream);
            }

        }

    }

    private Map<String, String> loadClientData() throws IOException {
        Map<String, String> ClientNameMap = new HashMap<>();
        try {
            FileInputStream clientInput = new FileInputStream(grid);
            Workbook clientWorkbook = new XSSFWorkbook(clientInput);
            Sheet clientSheet = clientWorkbook.getSheetAt(0);

            for (Row row : clientSheet) {
                Cell Appname = row.getCell(1);
                Cell client = row.getCell(2);
                if (Appname != null && client != null) {
                    String AppName = getCellValueAsString(Appname);
                    String clientName = getCellValueAsString(client);

                    if (!AppName.isEmpty()) {
                        ClientNameMap.put(AppName, clientName);
                    }
                } else {
                    System.out.println();
                }
            }
        } catch (IOException e) {
            System.err.println("error reading file " + e.getMessage());
        }
        return ClientNameMap;

    }

    private Map<String, String> loadEnglishNameData() throws IOException {
        Map<String, String> EnglishNameMap = new HashMap<>();
        try {
            FileInputStream clientInput = new FileInputStream(grid);
            Workbook clientWorkbook = new XSSFWorkbook(clientInput);
            Sheet clientSheet = clientWorkbook.getSheetAt(0);

            for (Row row : clientSheet) {
                Cell appname = row.getCell(1);
                Cell english = row.getCell(3);
                if (appname != null && english != null) {
                    String appName = getCellValueAsString(appname);
                    String englishName = getCellValueAsString(english);

                    if (!appName.isEmpty()) {
                        EnglishNameMap.put(appName, englishName);
                    }
                } else {
                    System.out.println();
                }
            }
        } catch (IOException e) {
            System.err.println("error reading file " + e.getMessage());
        }
        return EnglishNameMap;

    }

    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
            default:
                return "";
        }
    }




    public Map<String, Map<String, String[]>> preprocessSubtasks(String inputFilePath) throws IOException {
        Map<String, Map<String, String[]>> subtaskMap = new HashMap<>();
        try (FileInputStream inputStream = new FileInputStream(inputFilePath);
             Workbook workbook = new XSSFWorkbook(inputStream)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                String issueType = getCellValueAsString(row.getCell(1));
                String id = getCellValueAsString(row.getCell(0));
                String linkedId = getCellValueAsString(row.getCell(12));
                if (!linkedId.isEmpty() && !issueType.isEmpty()) {
                    String[] data = new String[]{
                            id,
                            getCellValueAsString(row.getCell(4)),  // status
                            getCellValueAsString(row.getCell(8)),  // assignee
                            getCellValueAsString(row.getCell(5)),  // created
                            getCellValueAsString(row.getCell(6))   // updated
                    };
                    subtaskMap
                            .computeIfAbsent(linkedId, k -> new HashMap<>())
                            .put(issueType, data);
                }
            }
        }
        return subtaskMap;
    }

    public String checkFeasibleStatus(String status){
        if (status.equalsIgnoreCase("Rejected")){
            return "Rejected";
        } else if (status.equalsIgnoreCase("Resolved")||status.equalsIgnoreCase("Release")) {
            return "Resolved";
        } else if (status.equalsIgnoreCase("Awaiting Client Response")) {
            return "ACR";
        } else if (status.equalsIgnoreCase("New")||status.equalsIgnoreCase("TODO")||status.equalsIgnoreCase("In Progress")||status.equalsIgnoreCase("Under Analysis")
                ||status.equalsIgnoreCase("ReOpen")||status.equalsIgnoreCase("Scheduled")||status.equalsIgnoreCase("Pending Release")||status.equalsIgnoreCase("Accepted in Roadmap")) {
            return "Active";
        }else if (status.equalsIgnoreCase("Closed")) {
            return "Closed";
        }
        return "";

    }

    private boolean isNullOrEmpty(String s) {
        return s == null || s.trim().isEmpty();
    }

    private void writeSummarySheet(Workbook wb) {
        Map<String,Integer> L2_com_Count = new LinkedHashMap<>();
        Map<String,Integer> L2_count = new LinkedHashMap<>();
        Map<String,Integer> L2_L3_Count = new LinkedHashMap<>();
        Map<String,Integer> L2_PMCount = new LinkedHashMap<>();
        Map<String,Integer> L2_BUG_Count = new LinkedHashMap<>();
        Map<String,Integer> L1rejectedcount = new LinkedHashMap<>();
        Map<String,Integer> L3Count = new LinkedHashMap<>();

        List<String> components = Arrays.asList("vConnect", "vDesigner","vDesigner 2.0","vFlow", "vFlow 2.0");

        for (String[] data : collectedIDs) {
            String comp = data[11];
            String l2 = data[30];

            if (!isNullOrEmpty(l2) && components.contains(comp)) {
                L2_com_Count.put(comp, L2_com_Count.getOrDefault(comp, 0) + 1);

                if (!isNullOrEmpty(data[35])) {
                    L2_L3_Count.put(comp, L2_L3_Count.getOrDefault(comp, 0) + 1);
                }

                if (isNullOrEmpty(data[35])) {
                    if (!isNullOrEmpty(data[45])) {
                        L2_PMCount.put(comp, L2_PMCount.getOrDefault(comp, 0) + 1);
                    } else if (!isNullOrEmpty(data[40])) {
                        L2_BUG_Count.put(comp, L2_BUG_Count.getOrDefault(comp, 0) + 1);
                    }
                }

                if (isNullOrEmpty(data[35]) && isNullOrEmpty(data[40]) && isNullOrEmpty(data[45])) {
                    L2_count.put(comp, L2_count.getOrDefault(comp, 0) + 1);
                }
            }

            if ("Rejected".equalsIgnoreCase(data[16]) || "ReOpen".equalsIgnoreCase(data[16]) && components.contains(comp) &&
                    isNullOrEmpty(data[30]) && isNullOrEmpty(data[35]) &&
                    isNullOrEmpty(data[40]) && isNullOrEmpty(data[45])) {
                L1rejectedcount.put(comp, L1rejectedcount.getOrDefault(comp, 0) + 1);
            }

            if (components.contains(comp) && isNullOrEmpty(l2) && !isNullOrEmpty(data[35])) {
                L3Count.put(comp, L3Count.getOrDefault(comp, 0) + 1);
            }
        }

        Sheet summary = wb.createSheet("Pivot");


        // Create a bold font and colored style for headers
        CellStyle headerStyle = wb.createCellStyle();
        Font headerFont = wb.createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setAlignment(HorizontalAlignment.RIGHT);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // Header Row
        Row h = summary.createRow(0);
        String[] headers = {"Components", "Cumulative L2", "L2", "L2-L3", "L2-PM", "L2-BUG", "L3", "L1 Rejected", "Grand Total"};
        for (int i = 0; i < headers.length; i++) {
            Cell cell = h.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
            summary.setColumnWidth(i, 3000);
        }

        int r = 1;
        int totalL2 = 0, totalL2Only = 0, totalL2_L3 = 0, totalPM = 0, totalBug = 0, totalL3 = 0, totalRejected = 0;

        for (String comp : components) {
            int l2Total = L2_com_Count.getOrDefault(comp, 0);
            int l2Only = L2_count.getOrDefault(comp, 0);
            int l2L3 = L2_L3_Count.getOrDefault(comp, 0);
            int l2PM = L2_PMCount.getOrDefault(comp, 0);
            int l2Bug = L2_BUG_Count.getOrDefault(comp, 0);
            int l3 = L3Count.getOrDefault(comp, 0);
            int rejected = L1rejectedcount.getOrDefault(comp, 0);

            Row row = summary.createRow(r++);
            row.createCell(0).setCellValue(comp);
            row.createCell(1).setCellValue(l2Total);
            row.createCell(2).setCellValue(l2Only);
            row.createCell(3).setCellValue(l2L3);
            row.createCell(4).setCellValue(l2PM);
            row.createCell(5).setCellValue(l2Bug);
            row.createCell(6).setCellValue(l3);
            row.createCell(7).setCellValue(rejected);
            row.createCell(8).setCellValue(l2Total + l3 + rejected);

            totalL2 += l2Total;
            totalL2Only += l2Only;
            totalL2_L3 += l2L3;
            totalPM += l2PM;
            totalBug += l2Bug;
            totalL3 += l3;
            totalRejected += rejected;
        }

        // Grand Total row
        Row totalRow = summary.createRow(r);
        String[] totalValues = {"Grand Total",
                String.valueOf(totalL2), String.valueOf(totalL2Only), String.valueOf(totalL2_L3),
                String.valueOf(totalPM), String.valueOf(totalBug), String.valueOf(totalL3),
                String.valueOf(totalRejected), String.valueOf(totalL2 + totalL3 + totalRejected)
        };

        for (int i = 0; i < totalValues.length; i++) {
            Cell cell = totalRow.createCell(i);
            cell.setCellValue(totalValues[i]);
            cell.setCellStyle(headerStyle);
        }
    }











}

