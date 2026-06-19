package com.bkademy.excelconverter.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class ExcelOutputFormatterTest {

    @Test
    void formatsHeaderAndAutoSizesColumns() {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Kết quả");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("STT");
            header.createCell(1).setCellValue("MSSV");
            header.createCell(2).setCellValue("Điểm rèn luyện");
            header.createCell(3).setCellValue("Kỳ học");

            Row row = sheet.createRow(1);
            row.createCell(0).setCellValue(1);
            row.createCell(1).setCellValue("1855010004");
            row.createCell(2).setCellValue("80");
            row.createCell(3).setCellValue("20232");

            ExcelOutputFormatter.format(sheet);

            Font headerFont = workbook.getFontAt(sheet.getRow(0).getCell(0).getCellStyle().getFontIndex());
            assertEquals("Times New Roman", headerFont.getFontName());
            assertTrue(headerFont.getBold());

            assertEquals(HorizontalAlignment.CENTER, sheet.getRow(0).getCell(0).getCellStyle().getAlignment());
            assertTrue(sheet.getColumnWidth(1) > sheet.getColumnWidth(0));
        } catch (Exception e) {
            fail(e);
        }
    }
}
