package org.example;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.util.*;


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


        List<Candidate> candidates = readCandidates(inputFilePath, sheetName);
        System.out.println("✅ Candidates Loaded:\n");


//        List of lists containing the rooms and who is in them
        List<List<Candidate>> lrwRooms = roomSorter(candidates);


        List<String> roomNames = List.of("Room 2", "Room 4", "Room 7", "Room 14");

        for (int i = 0; i < lrwRooms.size(); i++) {
            List<Candidate> room = lrwRooms.get(i);
            String roomName = roomNames.get(i);
            System.out.println(roomName + ":");

            for (Candidate c : room) {
                System.out.println("  " + c.getFirstName() + " " + c.getLastName() + " - " + c.getCandidateNumber());
            }

            System.out.println(); // empty line between rooms
        }


    }









//    Method for reading the Excel spreadsheet
public static List<Candidate> readCandidates(String filePath, String sheetName) throws Exception {
    Map<String, Candidate> uniqueCandidates = new LinkedHashMap<>();

    try (FileInputStream fis = new FileInputStream(filePath);
         Workbook workbook = new XSSFWorkbook(fis)) {

        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) throw new IllegalArgumentException("Sheet not found: " + sheetName);

        for (int i = 1; i <= sheet.getLastRowNum(); i++) { // skip header row
            Row row = sheet.getRow(i);
            if (row == null) continue;

            String candidateNumber = getCellValueAsString(row.getCell(0));
            String firstName = getCellValueAsString(row.getCell(1));
            String lastName = getCellValueAsString(row.getCell(2));
            String profession = getCellValueAsString(row.getCell(10));

            if (candidateNumber.isEmpty()) continue;

            if (uniqueCandidates.containsKey(candidateNumber)) {
                Candidate existing = uniqueCandidates.get(candidateNumber);

                // Fill in missing fields if new row has better data
                if ((existing.getFirstName() == null || existing.getFirstName().isBlank()) && !firstName.isBlank()) {
                    existing.setFirstName(firstName);
                }

                if ((existing.getLastName() == null || existing.getLastName().isBlank()) && !lastName.isBlank()) {
                    existing.setLastName(lastName);
                }

                if ((existing.getProfession() == null || existing.getProfession().isBlank()) && !profession.isBlank()) {
                    existing.setProfession(profession);
                }

            } else {
                uniqueCandidates.put(candidateNumber, new Candidate(candidateNumber, firstName, lastName, profession));
            }
        }
    }

    return new ArrayList<>(uniqueCandidates.values());
}



    //        method to sort candidates into their respective room numbers for LRW
    public static List<List<Candidate>> roomSorter(List<Candidate> candidates) {
        List<List<Candidate>> rooms = new ArrayList<>();
        rooms.add(new ArrayList<>()); // Room 2
        rooms.add(new ArrayList<>()); // Room 4
        rooms.add(new ArrayList<>()); // Room 7
        rooms.add(new ArrayList<>()); // Room 14

        int roomCapacity = 15;
        int currentRoom = 0;

        for (Candidate c : candidates) {
            boolean assigned = false;

            // Try to assign to one of the first 3 rooms
            for (int i = 0; i < 3; i++) {
                int index = (currentRoom + i) % 3;
                if (rooms.get(index).size() < roomCapacity) {
                    rooms.get(index).add(c);
                    currentRoom = (index + 1) % 3;  // rotate only after successful assign
                    assigned = true;
                    break;
                }
            }

            // Fallback to Room 14 if all 3 are full
            if (!assigned) {
                rooms.get(3).add(c);
            }
        }

        System.out.println("Room 2: " + rooms.get(0).size());
        System.out.println("Room 4: " + rooms.get(1).size());
        System.out.println("Room 7: " + rooms.get(2).size());
        System.out.println("Room 14: " + rooms.get(3).size());

        return rooms;
    }


    //        Helper function to convert spreadsheet values to Strings
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
