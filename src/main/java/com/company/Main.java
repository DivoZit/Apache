package com.company;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

public class Main {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        createExcelFile();

        File file = new File("excel.xlsx");
        Workbook wb = WorkbookFactory.create(file);
        Sheet s1 = wb.getSheetAt(0);
        int lastRowIndex = s1.getLastRowNum();
        for (int i = 1; i <= lastRowIndex; i++) {
            Row row = s1.getRow(i);
            System.out.printf("%-10s|%10d|%10s\n", row.getCell(0).getStringCellValue(),
                    (long) row.getCell(1).getNumericCellValue(),
                    row.getCell(2).getStringCellValue());
        }

    }

    private static void createExcelFile() throws IOException {
        List<String> animals = Arrays.asList("Zebra", "Giraffe", "Wolf", "Lion", "Otter", "Sea Lion");
        List<Integer> ages = Arrays.asList(2, 4, 5, 1, 1, 2);
        Workbook wb = new XSSFWorkbook();
        CreationHelper ch = wb.getCreationHelper();
        Sheet s1 = wb.createSheet("First sheet");
        Row r1 = s1.createRow(0);
        RichTextString rts = ch.createRichTextString("Animal");
        Font font = wb.createFont();
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        rts.applyFont(font);
        r1.createCell(0).setCellValue(rts);
        rts = ch.createRichTextString("Age");
        rts.applyFont(font);
        r1.createCell(1).setCellValue(rts);
        int rowCount = 1;
        int i = 0;
        for (String animal : animals) {
            Row row = s1.createRow(rowCount++);
            row.createCell(0).setCellValue(animal);
            row.createCell(1).setCellValue(ages.get(i++));
            row.createCell(2).setCellValue("Trecias");
        }
        FileOutputStream fileOutputStream = new FileOutputStream("excel.xlsx");
        wb.write(fileOutputStream);
    }
}
