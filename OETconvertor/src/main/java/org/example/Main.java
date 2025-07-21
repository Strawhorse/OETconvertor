package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;


//TIP To <b>Run</b> code, press <shortcut actionId="Run"/> or
// click the <icon src="AllIcons.Actions.Execute"/> icon in the gutter.
public class Main {
    public static void main(String[] args) throws Exception {
        //TIP Press <shortcut actionId="ShowIntentionActions"/> with your caret at the highlighted text
        // to see how IntelliJ IDEA suggests fixing it.

        String inputFilePath = "";
        String sheetName = "LRW";

        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Select Excel File you wish to import");
        int result = fileChooser.showOpenDialog(null);

        if (result == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            inputFilePath = selectedFile.getAbsolutePath();

        } else {
            System.out.println("File selection cancelled.");
        }


        List<List<Candidate>> candidates = ExcelReader.readCandidates(inputFilePath, sheetName);

        System.out.println("✅ Candidates Loaded:");
        for (List<Candidate> c : candidates) {
            System.out.println(c.toString());
        }

    }





//    Method for reading the Excel spreadsheet
    public static class ExcelReader {

        public static List<List<Candidate>> readCandidates(String filePath, String sheetName) throws Exception {
            List<Candidate> listeningCandidates = new ArrayList<>();
            List<Candidate> readingCandidates = new ArrayList<>();
            List<Candidate> writingCandidates = new ArrayList<>();

            try (FileInputStream fis = new FileInputStream(filePath);
                 Workbook workbook = new XSSFWorkbook(fis)) {

                Sheet sheet = workbook.getSheet(sheetName);
                if (sheet == null) throw new IllegalArgumentException("Sheet not found: " + sheetName);

                for (int i = 1; i <= sheet.getLastRowNum(); i++) { // skip header row
                    Row row = sheet.getRow(i);
                    if (row == null) continue;

                    String testType = getCellValueAsString(row.getCell(6));
                    String candidateNumber = getCellValueAsString(row.getCell(0));
                    String firstName = getCellValueAsString(row.getCell(1));
                    String lastName = getCellValueAsString(row.getCell(2));
                    String profession = getCellValueAsString(row.getCell(10));

                    if (testType.contains("Listening")) {
                        listeningCandidates.add(new Candidate(candidateNumber, firstName, lastName, profession));
                    } else if (testType.contains("Reading")) {
                        readingCandidates.add(new Candidate(candidateNumber, firstName, lastName, profession));
                    } else {
                    writingCandidates.add(new Candidate(candidateNumber, firstName, lastName, profession));
                    }
            }


            }
            List<List<Candidate>> masterList = new ArrayList<>();
            masterList.add(listeningCandidates);
            masterList.add(readingCandidates);
            masterList.add(writingCandidates);

            System.out.println("Outputting listening, reading, and writing lists...");
            return masterList;
        }

        private static String getCellValueAsString(Cell cell) {
            if (cell == null) return "";

            return switch (cell.getCellType()) {
                case STRING -> cell.getStringCellValue();
                case NUMERIC -> String.valueOf((long) cell.getNumericCellValue());
                case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
                case FORMULA -> cell.getCellFormula();
                default -> "";
            };
        }
    }



//    method to write to the new spreadsheet

}
