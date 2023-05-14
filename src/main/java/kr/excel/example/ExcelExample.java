package kr.excel.example;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ExcelExample {
    public static void main(String[] args) {
        try {
            FileInputStream file = new FileInputStream(new File("example.xlsx"));
            Workbook workbook = WorkbookFactory.create(file);
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                for (Cell cell : row) {
                    switch (cell.getCellType()) {
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                Date dateCellValue = cell.getDateCellValue();
                                DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-yy");
                                String format = dateFormat.format(dateCellValue);
                                System.out.print(format + "\t");
                            } else {
                                double numericValue = cell.getNumericCellValue();
                                if (numericValue == Math.floor(numericValue)) {
                                    int intValue = (int) numericValue;
                                    System.out.print(intValue + "\t");
                                } else {
                                    System.out.print(numericValue + "\t");
                                }

                            }
                            break;
                        case STRING:
                            String stringCellValue = cell.getStringCellValue();
                            System.out.print(stringCellValue + "\t");
                            break;
                        case BOOLEAN:
                            boolean booleanCellValue = cell.getBooleanCellValue();
                            System.out.print(booleanCellValue + "\t");
                            break;
                        case FORMULA:
                            String cellFormula = cell.getCellFormula();
                            System.out.print(cellFormula + "\t");
                            break;
                        case BLANK:
                            System.out.print("\t");
                            break;
                        default:
                            System.out.print("\t");
                            break;
                    }
                }
                System.out.println();
            }
            file.close();
            System.out.println("엑셀에서 데이터 읽어오기 성공");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}