import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Coalition {

    public static void main(String[] args) {
        File resourcesDirectory = new File("src/main/resources");
        File[] files = resourcesDirectory.listFiles((dir, name) -> name.endsWith(".xlsx")/* || name.endsWith(".xls") || name.endsWith(".csv")*/);

        if (files != null) {
            // Create a master workbook
            Workbook masterWorkbook = new XSSFWorkbook();

            for (File file : files) {
                if (!(file.getName().contains("Master"))) {
                    System.out.println("Reading file: " + file.getName());
                    //readExcel(file);
                    readAndCopyColumnASingleFile(file, masterWorkbook);
                }
            }
            // Save the master workbook
            try (FileOutputStream fos = new FileOutputStream("src/main/resources/MasterWorkbook.xlsx")) {
                masterWorkbook.write(fos);
                System.out.println("Master workbook created successfully.");
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            System.out.println("No Excel files found in the resources directory.");
        }
    }



    private static void readAndCopyColumnASingleFile(File file, Workbook masterWorkbook) {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = WorkbookFactory.create(fis)) {
            String storeName = getNumberFromFileName(file.getName());
                    //.replaceAll(".xlsx", "");
            //storeName = storeName.substring(storeName.length() - 4);

            String sheetName = "Consolidated_Data";

            // Get or create the master sheet
            Sheet masterSheet = masterWorkbook.getSheet(sheetName);
            if (masterSheet == null) {
                masterSheet = masterWorkbook.createSheet(sheetName);
            }

            // Get the last row index in the master sheet
            int lastRowIndex = masterSheet.getLastRowNum();
            if (lastRowIndex < 0)
                lastRowIndex = 0;

            // Iterate through rows and copy data from column A to the master sheet
            int rowCount = 0;
            for (Row row : workbook.getSheetAt(0)) {
                Row masterRow = masterSheet.createRow(lastRowIndex + rowCount++);

                // Copy data from column A
                if (rowCount != 1) {
                    Cell masterCellA = masterRow.createCell(0, CellType.STRING);
                    masterCellA.setCellValue(storeName);
                }

                // Copy data from column A to column B
                Cell cellA = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                Cell masterCellB = masterRow.createCell(1, CellType.STRING);
                masterCellB.setCellValue(cellA.toString());

                // Copy data from column E to column C
                Cell cellE = row.getCell(4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                Cell masterCellC = masterRow.createCell(2, CellType.STRING);
                masterCellC.setCellValue(cellE.toString());
            }

            // Update the last row index in the master sheet
            lastRowIndex += rowCount;

        } catch (IOException | EncryptedDocumentException e) {
            e.printStackTrace();
        }
    }

    private static void readExcel(File file) {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = WorkbookFactory.create(fis)) {

            // Assuming there is only one sheet in each Excel file
            Sheet sheet = workbook.getSheetAt(0);

            // Iterate through rows and columns
            for (Row row : sheet) {
                for (Cell cell : row) {
                    System.out.print(cell.toString() + "\t");
                }
                System.out.println();
            }

        } catch (IOException | EncryptedDocumentException e) {
            e.printStackTrace();
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
