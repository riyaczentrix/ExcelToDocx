import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

public class EmailProcessorXLSX {

    public static class EmailEntry {
        private String docketNo;
        private String mailListId;
        private String mailId;
        private String ticketId;
        private String dispositionName;
        private String subDispositionName;
        private String priorityName;
        private String problemReported;
        private String assignedToDeptName;
        private String processedBodyCleaned;
        private String solution;
        private int cluster;

        public EmailEntry(String docketNo, String mailListId, String mailId, String ticketId,
                          String dispositionName, String subDispositionName, String priorityName,
                          String problemReported, String assignedToDeptName, String processedBodyCleaned,
                          String solution, int cluster) {
            this.docketNo = docketNo;
            this.mailListId = mailListId;
            this.mailId = mailId;
            this.ticketId = ticketId;
            this.dispositionName = dispositionName;
            this.subDispositionName = subDispositionName;
            this.priorityName = priorityName;
            this.problemReported = problemReported;
            this.assignedToDeptName = assignedToDeptName;
            this.processedBodyCleaned = processedBodyCleaned;
            this.solution = solution;
            this.cluster = cluster;
        }

        public String getDocketNo() { return docketNo; }
        public String getMailListId() { return mailListId; }
        public String getMailId() { return mailId; }
        public String getTicketId() { return ticketId; }
        public String getDispositionName() { return dispositionName; }
        public String getSubDispositionName() { return subDispositionName; }
        public String getPriorityName() { return priorityName; }
        public String getProblemReported() { return problemReported; }
        public String getAssignedToDeptName() { return assignedToDeptName; }
        public String getProcessedBodyCleaned() { return processedBodyCleaned; }
        public String getSolution() { return solution; }
        public int getCluster() { return cluster; }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;
            EmailEntry that = (EmailEntry) o;
            return cluster == that.cluster &&
                    Objects.equals(docketNo, that.docketNo) &&
                    Objects.equals(mailListId, that.mailListId) &&
                    Objects.equals(mailId, that.mailId) &&
                    Objects.equals(ticketId, that.ticketId) &&
                    Objects.equals(dispositionName, that.dispositionName) &&
                    Objects.equals(subDispositionName, that.subDispositionName) &&
                    Objects.equals(priorityName, that.priorityName) &&
                    Objects.equals(problemReported, that.problemReported) &&
                    Objects.equals(assignedToDeptName, that.assignedToDeptName) &&
                    Objects.equals(processedBodyCleaned, that.processedBodyCleaned) &&
                    Objects.equals(solution, that.solution);
        }

        @Override
        public int hashCode() {
            return Objects.hash(docketNo, mailListId, mailId, ticketId, dispositionName,
                    subDispositionName, priorityName, problemReported, assignedToDeptName,
                    processedBodyCleaned, solution, cluster);
        }
    }

    /**
     * Processes clustered email data and writes directly to DOCX to minimize memory usage.
     */
    public static void processClusteredEmailsAndSaveToDocx(InputStream inputStream, String outputPath) throws IOException {
        try (XWPFDocument document = new XWPFDocument()) {
            // Add page numbers to the footer
            XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(document);
            XWPFFooter footer = headerFooterPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);
            XWPFParagraph footerParagraph = footer.createParagraph();
            footerParagraph.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun footerRun = footerParagraph.createRun();
            footerRun.setFontFamily("Consolas");
            footerRun.setFontSize(10);
            CTFldChar fldCharBegin = footerRun.getCTR().addNewFldChar();
            fldCharBegin.setFldCharType(STFldCharType.BEGIN);
            footerRun.getCTR().addNewInstrText().setStringValue("PAGE");
            CTFldChar fldCharEnd = footerRun.getCTR().addNewFldChar();
            fldCharEnd.setFldCharType(STFldCharType.END);

            Map<Integer, List<EmailEntry>> groupedClusters = new TreeMap<>();

            // Read XLSX and group clusters
            try (Workbook workbook = new XSSFWorkbook(inputStream)) {
                Sheet sheet = workbook.getSheetAt(0);
                if (sheet == null) {
                    throw new IOException("No sheets found in the XLSX workbook.");
                }

                Row headerRow = sheet.getRow(0);
                if (headerRow == null) {
                    throw new IOException("XLSX data is empty or header row is missing.");
                }

                Map<String, Integer> headerMap = new HashMap<>();
                for (Cell cell : headerRow) {
                    headerMap.put(cell.getStringCellValue().trim(), cell.getColumnIndex());
                }

                int docketNoIdx = headerMap.getOrDefault("docket_no", -1);
                int mailListIdIdx = headerMap.getOrDefault("mail_list_id", -1);
                int mailIdIdx = headerMap.getOrDefault("mail_id", -1);
                int ticketIdIdx = headerMap.getOrDefault("ticket_id", -1);
                int dispositionNameIdx = headerMap.getOrDefault("disposition_name", -1);
                int subDispositionNameIdx = headerMap.getOrDefault("sub_disposition_name", -1);
                int priorityNameIdx = headerMap.getOrDefault("priority_name", -1);
                int problemReportedIdx = headerMap.getOrDefault("problem_reported", -1);
                int assignedToDeptNameIdx = headerMap.getOrDefault("assigned_to_dept_name", -1);
                int processedBodyCleanedIdx = headerMap.getOrDefault("ProcessedBody_cleaned", -1);
                int solutionIdx = headerMap.getOrDefault("Solution", -1);
                int clusterIdx = headerMap.getOrDefault("Cluster", -1);

                if (docketNoIdx == -1 || mailListIdIdx == -1 || mailIdIdx == -1 || ticketIdIdx == -1 ||
                        dispositionNameIdx == -1 || subDispositionNameIdx == -1 || priorityNameIdx == -1 ||
                        problemReportedIdx == -1 || assignedToDeptNameIdx == -1 ||
                        processedBodyCleanedIdx == -1 || solutionIdx == -1 || clusterIdx == -1) {
                    throw new IOException("Missing one or more required columns in XLSX.");
                }

                for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                    Row row = sheet.getRow(r);
                    if (row == null) {
                        continue;
                    }

                    try {
                        String docketNo = getCellValueAsString(row.getCell(docketNoIdx));
                        String mailListId = getCellValueAsString(row.getCell(mailListIdIdx));
                        String mailId = getCellValueAsString(row.getCell(mailIdIdx));
                        String ticketId = getCellValueAsString(row.getCell(ticketIdIdx));
                        String dispositionName = getCellValueAsString(row.getCell(dispositionNameIdx));
                        String subDispositionName = getCellValueAsString(row.getCell(subDispositionNameIdx));
                        String priorityName = getCellValueAsString(row.getCell(priorityNameIdx));
                        String problemReported = getCellValueAsString(row.getCell(problemReportedIdx));
                        String assignedToDeptName = getCellValueAsString(row.getCell(assignedToDeptNameIdx));
                        String processedBodyCleaned = getCellValueAsString(row.getCell(processedBodyCleanedIdx));
                        String solution = getCellValueAsString(row.getCell(solutionIdx));

                        Cell clusterCell = row.getCell(clusterIdx);
                        int cluster = -1;
                        if (clusterCell != null && clusterCell.getCellType() == CellType.NUMERIC) {
                            cluster = (int) clusterCell.getNumericCellValue();
                        } else {
                            System.err.println("Warning: Skipping row " + (r + 1) + " due to invalid or missing cluster number format: " + getCellValueAsString(clusterCell));
                            continue;
                        }

                        groupedClusters.computeIfAbsent(cluster, k -> new ArrayList<>()).add(
                                new EmailEntry(docketNo, mailListId, mailId, ticketId,
                                        dispositionName, subDispositionName, priorityName,
                                        problemReported, assignedToDeptName, processedBodyCleaned,
                                        solution, cluster));
                    } catch (Exception e) {
                        System.err.println("Warning: Skipping row " + (r + 1) + " due to data parsing error: " + e.getMessage());
                        continue;
                    }
                }
            }

            // Process each cluster and write to DOCX incrementally
            int clusterCount = 0;
            int maxClustersPerFile = 5000; // Split into multiple files if needed
            int fileCounter = 1;
            XWPFDocument currentDocument = document;

            for (Map.Entry<Integer, List<EmailEntry>> entry : groupedClusters.entrySet()) {
                clusterCount++;
                int clusterNumber = entry.getKey();
                List<EmailEntry> clusterRows = entry.getValue();
                System.out.println("Cluster " + clusterNumber + ": " + clusterRows.size() + " rows");

                // Start a new document if cluster count exceeds maxClustersPerFile
                if (clusterCount % maxClustersPerFile == 1 && clusterCount > 1) {
                    String splitOutputPath = outputPath.replace(".docx", "_" + fileCounter + ".docx");
                    try (FileOutputStream out = new FileOutputStream(splitOutputPath)) {
                        currentDocument.write(out);
                        System.out.println("DOCX file written to: " + splitOutputPath);
                    }
                    fileCounter++;
                    currentDocument = new XWPFDocument();
                    // Re-add footer to new document
                    headerFooterPolicy = new XWPFHeaderFooterPolicy(currentDocument);
                    footer = headerFooterPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);
                    footerParagraph = footer.createParagraph();
                    footerParagraph.setAlignment(ParagraphAlignment.CENTER);
                    footerRun = footerParagraph.createRun();
                    footerRun.setFontFamily("Consolas");
                    footerRun.setFontSize(10);
                    fldCharBegin = footerRun.getCTR().addNewFldChar();
                    fldCharBegin.setFldCharType(STFldCharType.BEGIN);
                    footerRun.getCTR().addNewInstrText().setStringValue("PAGE");
                    fldCharEnd = footerRun.getCTR().addNewFldChar();
                    fldCharEnd.setFldCharType(STFldCharType.END);
                }

                XWPFParagraph paragraph = currentDocument.createParagraph();
                List<String> lines;

                if (clusterNumber == 50000) {
                    // Treat every issue as a main issue for cluster 50000
                    for (EmailEntry mainIssue : clusterRows) {
                        lines = new ArrayList<>();
                        StringBuilder problemDetailsContent = new StringBuilder();
                        if (mainIssue.getProcessedBodyCleaned() != null && !mainIssue.getProcessedBodyCleaned().isEmpty()) {
                            // Truncate large fields to prevent memory issues
                            String processedBody = mainIssue.getProcessedBodyCleaned();
                            if (processedBody.length() > 10000) {
                                processedBody = processedBody.substring(0, 10000) + "... [Truncated]";
                            }
                            problemDetailsContent.append(processedBody);
                        }

                        lines.add("**Main Issue:**");
                        lines.add(mainIssue.getProblemReported());
                        lines.add("");
                        lines.add("Problem Details:");
                        lines.add(problemDetailsContent.toString());
                        lines.add("Disposition: " + mainIssue.getDispositionName());
                        lines.add("SubDisposition: " + mainIssue.getSubDispositionName());
                        lines.add("Priority: " + mainIssue.getPriorityName());
                        lines.add("- Docket No: " + mainIssue.getDocketNo());
                        lines.add("- Mail List ID: " + mainIssue.getMailListId());
                        lines.add("- Mail ID: " + mainIssue.getMailId());
                        lines.add("- Ticket ID: " + mainIssue.getTicketId());
                        lines.add("- Assigned To Dept: " + mainIssue.getAssignedToDeptName());
                        lines.add("");
                        lines.add("Solution:");
                        String solution = mainIssue.getSolution();
                        if (solution.length() > 10000) {
                            solution = solution.substring(0, 10000) + "... [Truncated]";
                        }
                        lines.add(solution);
                        lines.add("");
                        lines.add("**Similar issues:**");
                        lines.add("None");
                        lines.add("");

                        // Write lines to paragraph
                        for (String line : lines) {
                            XWPFRun run = paragraph.createRun();
                            run.setFontFamily("Consolas");
                            run.setFontSize(10);
                            if (line.startsWith("**Main Issue:**") || line.startsWith("**Similar issues:**")) {
                                run.setBold(true);
                                run.setText(line.replace("**", ""));
                            } else {
                                run.setText(line);
                            }
                            run.addBreak();
                        }

                        // Add separator within the same document
                        if (clusterCount < groupedClusters.size() || !clusterRows.isEmpty()) {
                            XWPFParagraph separator = currentDocument.createParagraph();
                            XWPFRun separatorRun = separator.createRun();
                            separatorRun.setFontFamily("Consolas");
                            separatorRun.setFontSize(10);
                            separatorRun.setText("---");
                            separatorRun.addBreak();
                        }
                    }
                } else {
                    // Original logic for other clusters
                    EmailEntry mainIssue = null;
                    List<EmailEntry> similarIssues = new ArrayList<>();

                    for (EmailEntry row : clusterRows) {
                        if (row.getSolution() != null && !row.getSolution().isEmpty()) {
                            mainIssue = row;
                            break;
                        }
                    }

                    if (mainIssue == null) {
                        if (!clusterRows.isEmpty()) {
                            mainIssue = clusterRows.get(0);
                            similarIssues.addAll(clusterRows.subList(1, clusterRows.size()));
                        } else {
                            continue;
                        }
                    } else {
                        for (EmailEntry row : clusterRows) {
                            if (!row.equals(mainIssue)) {
                                similarIssues.add(row);
                            }
                        }
                    }

                    StringBuilder problemDetailsContent = new StringBuilder();
                    if (mainIssue.getProcessedBodyCleaned() != null && !mainIssue.getProcessedBodyCleaned().isEmpty()) {
                        String processedBody = mainIssue.getProcessedBodyCleaned();
                        if (processedBody.length() > 10000) {
                            processedBody = processedBody.substring(0, 10000) + "... [Truncated]";
                        }
                        problemDetailsContent.append(processedBody);
                    }

                    lines = new ArrayList<>();
                    lines.add("**Main Issue:**");
                    lines.add(mainIssue.getProblemReported());
                    lines.add("");
                    lines.add("Problem Details:");
                    lines.add(problemDetailsContent.toString());
                    lines.add("Disposition: " + mainIssue.getDispositionName());
                    lines.add("SubDisposition: " + mainIssue.getSubDispositionName());
                    lines.add("Priority: " + mainIssue.getPriorityName());
                    lines.add("- Docket No: " + mainIssue.getDocketNo());
                    lines.add("- Mail List ID: " + mainIssue.getMailListId());
                    lines.add("- Mail ID: " + mainIssue.getMailId());
                    lines.add("- Ticket ID: " + mainIssue.getTicketId());
                    lines.add("- Assigned To Dept: " + mainIssue.getAssignedToDeptName());
                    lines.add("");
                    lines.add("Solution:");
                    String solution = mainIssue.getSolution();
                    if (solution.length() > 10000) {
                        solution = solution.substring(0, 10000) + "... [Truncated]";
                    }
                    lines.add(solution);
                    lines.add("");
                    lines.add("**Similar issues:**");

                    for (int i = 0; i < similarIssues.size(); i++) {
                        EmailEntry similarIssue = similarIssues.get(i);
                        lines.add("  " + (i + 1) + ". Issue reported : " + similarIssue.getProblemReported());
                        String processedBody = similarIssue.getProcessedBodyCleaned();
                        if (processedBody.length() > 10000) {
                            processedBody = processedBody.substring(0, 10000) + "... [Truncated]";
                        }
                        lines.add("     Problem Details: " + processedBody);
                        lines.add("     - Docket No: " + similarIssue.getDocketNo());
                        lines.add("     - Mail List ID: " + similarIssue.getMailListId());
                        lines.add("     - Mail ID: " + similarIssue.getMailId());
                        lines.add("     - Ticket ID: " + similarIssue.getTicketId());
                        lines.add("     - Disposition: " + similarIssue.getDispositionName());
                        lines.add("     - SubDisposition: " + similarIssue.getSubDispositionName());
                        lines.add("     - Priority: " + similarIssue.getPriorityName());
                        lines.add("     - Assigned To Dept: " + similarIssue.getAssignedToDeptName());
                        String similarSolution = similarIssue.getSolution();
                        if (similarSolution.length() > 10000) {
                            similarSolution = similarSolution.substring(0, 10000) + "... [Truncated]";
                        }
                        lines.add("     - Solution: " + similarSolution);
                        lines.add("");
                    }

                    // Write lines to paragraph
                    for (String line : lines) {
                        XWPFRun run = paragraph.createRun();
                        run.setFontFamily("Consolas");
                        run.setFontSize(10);
                        if (line.startsWith("**Main Issue:**") || line.startsWith("**Similar issues:**")) {
                            run.setBold(true);
                            run.setText(line.replace("**", ""));
                        } else {
                            run.setText(line);
                        }
                        run.addBreak();
                    }

                    // Add separator within the same document
                    if (clusterCount < groupedClusters.size()) {
                        XWPFParagraph separator = currentDocument.createParagraph();
                        XWPFRun separatorRun = separator.createRun();
                        separatorRun.setFontFamily("Consolas");
                        separatorRun.setFontSize(10);
                        separatorRun.setText("---");
                        separatorRun.addBreak();
                    }
                }

                // Periodically trigger garbage collection to free memory
                if (clusterCount % 1000 == 0) {
                    System.gc();
                    System.out.println("Memory used after cluster " + clusterNumber + ": " +
                            (Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory()) / (1024 * 1024) + " MB");
                }
            }

            // Write the final document
            String finalOutputPath = fileCounter == 1 ? outputPath : outputPath.replace(".docx", "_" + fileCounter + ".docx");
            try (FileOutputStream out = new FileOutputStream(finalOutputPath)) {
                currentDocument.write(out);
                System.out.println("DOCX file written to: " + finalOutputPath);
            }
        }
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell).trim();
    }

    public static void main(String[] args) {
        String filePath = args.length > 0 ? args[0] : "clustered_emails.xlsx";
        String outputDocxPath = "ClusteredEmailReport.docx";

        try (InputStream fis = new FileInputStream(new File(filePath))) {
            processClusteredEmailsAndSaveToDocx(fis, outputDocxPath);
            System.out.println("Report processing complete.");
        } catch (IOException e) {
            System.err.println("Error processing the file or saving DOCX: " + e.getMessage());
            e.printStackTrace();
        }
    }
}