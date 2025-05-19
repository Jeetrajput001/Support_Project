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
    public class DailyReport {


    private String inputFilePath = "/home/decimal/Downloads/DVES_Ticket_Report1905.xlsx";  //input file path
    private String outputFilePath = "/home/decimal/Support/Auto_Generated_Report/DailyReport/19-05-2025.xlsx"; //output file path
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
            Sheet newSheet = newWorkbook.createSheet("Daily Report");

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


    public String[] subtask(String linkID,String taskName){
        String [] result = new String[5];
        String SrId=linkID;
//        String type=taskName;
//        String id ="1";
//        String status="";
//        String asignee="";
//        String createddate="";
//        String updateDate="";
        try (FileInputStream inputStream = new FileInputStream(inputFilePath);
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            Sheet originalSheet = workbook.getSheetAt(0);
            for (Row row: originalSheet){
                if (row!=null) {
                    Cell issueTypeCell = row.getCell(1); // Assuming issue type is in the second column
                    Cell idCell = row.getCell(0); // Assuming ID is in the first column
                    Cell linkissued = row.getCell(12); // Assuming linkissued is in the 13th column
                    Cell statusCell=row.getCell(4);
                    Cell AsigneeCell=row.getCell(8);
                    Cell createDatecell = row.getCell(5);
                    Cell updateDatecell = row.getCell(6);
                    if (issueTypeCell != null && idCell != null && linkissued != null) {
                        String issueType = issueTypeCell.getStringCellValue();
                        String linked = linkissued.getStringCellValue();
                        if(issueType.equalsIgnoreCase(taskName) && linked.equalsIgnoreCase(SrId)){
                            result[0]=idCell.getStringCellValue();
                            result[1]=getCellValueAsString(statusCell);
                            result[2]=getCellValueAsString(AsigneeCell);
                            result[3] = getCellValueAsString(createDatecell);
                            result[4] = getCellValueAsString(updateDatecell);

                            break;

                        }
                    }

                }


            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        return result;

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

    private void writeSummarySheet(Workbook wb) {
        // 1) Build your two count‐maps from collectedIDs:
        Map<String,Integer> feasibleCount = new LinkedHashMap<>();
        Map<String,Integer> pmBugCount   = new LinkedHashMap<>();
        for (String[] data : collectedIDs) {
            String fs = data[5];  // feasible status
            String pm = data[6];  // active PM/BUG
            feasibleCount.put(fs, feasibleCount.getOrDefault(fs,0)+1);
            if (!pm.isEmpty()) {
                // we count under the same key so the rows line up by feasible-status
                pmBugCount.put(fs, pmBugCount.getOrDefault(fs,0)+1);
            }
        }

        // 2) Create the sheet
        Sheet summary = wb.createSheet("Pivot");

        // 3) Header row
        Row h = summary.createRow(0);
        h.createCell(0).setCellValue("Feasible Status");
        h.createCell(1).setCellValue("Feasible Status Count");
        h.createCell(2).setCellValue("Active PM/BUG Count");
        summary.setColumnWidth(0, 6000);
        summary.setColumnWidth(1, 6000);
        summary.setColumnWidth(2, 6000);

        // 4) Data rows
        int r = 1;
        int totalF = 0, totalPM = 0;
        for (Map.Entry<String,Integer> e : feasibleCount.entrySet()) {
            String fs = e.getKey();
            int   c1 = e.getValue();
            int   c2 = pmBugCount.getOrDefault(fs,0);
            Row row = summary.createRow(r++);
            row.createCell(0).setCellValue(fs);
            row.createCell(1).setCellValue(c1);
            row.createCell(2).setCellValue(c2);
            totalF += c1;
            totalPM+= c2;
        }

        // 5) Grand Total
        Row tot = summary.createRow(r);
        tot.createCell(0).setCellValue("Grand Total");
        tot.createCell(1).setCellValue(totalF);
        tot.createCell(2).setCellValue(totalPM);
    }









}

