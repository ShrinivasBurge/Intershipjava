package com.demo.InternProject;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;s
import java.util.*;

public class EmployeeAnalyzer {
    public static void main(String[] args) {
        String filePath = "C:/Users/HP/eclipse-workspace/InternProject/Assignment_Timecard.xlsx";
        try {
            analyzeExcelFile(filePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    private static void analyzeExcelFile(String filePath) throws IOException {
        Map<String, List<Shift>> employees = new HashMap<>();

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); 
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row != null) {
                    String name = row.getCell(7).getStringCellValue(); 
                    String timeIn = row.getCell(2).getStringCellValue(); 
                    String timeOut = row.getCell(3).getStringCellValue(); 

                    employees.computeIfAbsent(name, k -> new ArrayList<>())
                            .add(new Shift(timeIn, timeOut));
                }
            }
        }
        try (FileWriter writer = new FileWriter("output.txt")) {
            for (Map.Entry<String, List<Shift>> entry : employees.entrySet()) {
                String name = entry.getKey();
                List<Shift> shifts = entry.getValue();

                for (int i = 0; i < shifts.size() - 1; i++) {
                    if (daysBetween(shifts.get(i).getEndDate(), shifts.get(i + 1).getStartDate()) == 1) {
                        System.out.println(name + " has worked for 7 consecutive days.");
                        writer.write(name + " has worked for 7 consecutive days.\n");
                    }
                    long timeBetweenShifts = minutesBetween(shifts.get(i).getEndDate(), shifts.get(i + 1).getStartDate());
                    if (60 < timeBetweenShifts && timeBetweenShifts < 600) {
                        System.out.println(name + " has less than 10 hours between shifts but greater than 1 hour.");
                        writer.write(name + " has less than 10 hours between shifts but greater than 1 hour.\n");
                    }
                    if (shifts.get(i).getShiftDuration() > 840) {
                        System.out.println(name + " has worked for more than 14 hours in a single shift.");
                        writer.write(name + " has worked for more than 14 hours in a single shift.\n");
                    }
                }
            }
        }
    }
    private static class Shift {
        private final Date startDate;
        private final Date endDate;

        Shift(String timeIn, String timeOut) {
            if (timeIn == null || timeOut == null || timeIn.trim().isEmpty() || timeOut.trim().isEmpty()) {
                throw new RuntimeException("Date string is empty or null");
            }
            this.startDate = parseDate(timeIn);
            this.endDate = parseDate(timeOut);
        }
        public Date getStartDate() {
            return startDate;
        }
        public Date getEndDate() {
            return endDate;
        }
        public long getShiftDuration() {
            return (endDate.getTime() - startDate.getTime()) / (1000 * 60);
        }
        private Date parseDate(String dateStr) {
            SimpleDateFormat[] dateFormats = {
                    new SimpleDateFormat("MM-dd-yyyy hh:mm a"),
                    new SimpleDateFormat("MM-dd-yyyy hh:mma"),
                    new SimpleDateFormat("MM/dd/yyyy hh:mm a"),
                    new SimpleDateFormat("MM/dd/yyyy hh:mma")
            };
            for (SimpleDateFormat dateFormat : dateFormats) {
                try {
                    return dateFormat.parse(dateStr);
                } catch (ParseException ignored) {
                }
            }
            throw new RuntimeException("Error parsing date: " + dateStr);
        }
    }
    private static long daysBetween(Date firstDate, Date secondDate) {
        return (secondDate.getTime() - firstDate.getTime()) / (1000 * 60 * 60 * 24);
    }

    private static long minutesBetween(Date firstDate, Date secondDate) {
        return (secondDate.getTime() - firstDate.getTime()) / (1000 * 60);
    }
}
