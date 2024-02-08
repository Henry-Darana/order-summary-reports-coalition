import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class CheckValidFiles {

    public static void main(String[] args) {
        File resourcesDirectory = new File("src/main/resources");
        File[] files = resourcesDirectory.listFiles(
                (dir, name) -> name.endsWith(".xlsx") || name.endsWith(".xls") || name.endsWith(".csv"));

        System.out.println("File Name" + "\tAble To Open" + "\tHas Data");

        if (files != null) {

            for (File file : files) {
                if (!(file.getName().contains("Master"))) {
                    //System.out.println("Reading file: " + file.getName());
                    boolean ableToOpen = true;
                    boolean hasData = false;

                    try {
                        hasData = checkFile(file);
                    } catch (IOException e) {
                        ableToOpen = false;
                    }

                    System.out.println(getNumberFromFileName(file.getName()) + "\t" + ableToOpen + "\t" + hasData);

                }
            }
        } else {
            System.out.println("No Excel files found in the resources directory.");
        }
    }

    private static boolean checkFile(File file) throws IOException {
        try (FileInputStream fis = new FileInputStream(file)) {
            Workbook workbook = WorkbookFactory.create(fis);
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(1); // Assuming row 2 corresponds to index 1

            // Check if the row exists and if cell 1 in row 2 has data
            return row != null && row.getCell(0) != null && row.getCell(0).getCellType() != CellType.BLANK;
        }
    }

    private static String getNumberFromFileName1(String fileName) {
        // Regular expression pattern to match the numbers
        Pattern pattern = Pattern.compile("#(\\d+)");
        Matcher matcher = pattern.matcher(fileName);

        // Find and return the first number found in the file name
        if (matcher.find()) {
            return matcher.group();
        } else {
            return ""; // If no number found, return empty string
        }
    }

    public static String getNumberFromFileName(String fileName) {
        fileName = fileName.replace("(1)", "");
        int lastDotIndex = fileName.lastIndexOf('.');
        if (lastDotIndex != -1 && lastDotIndex >= 4) {
            return "IHOP #" + fileName.substring(lastDotIndex - 4, lastDotIndex);
        } else {
            // If there are less than 4 characters before the last dot, return an empty string
            return "";
        }
    }
}
