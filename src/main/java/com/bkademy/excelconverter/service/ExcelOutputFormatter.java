package com.bkademy.excelconverter.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

final class ExcelOutputFormatter {

    private static final String FONT_NAME = "Times New Roman";
    private static final short FONT_SIZE = 11;
    private static final int MIN_COLUMN_WIDTH = 2800;
    private static final int MAX_COLUMN_WIDTH = 18000;

    private ExcelOutputFormatter() {
    }

    static void format(Sheet sheet) {
        Workbook workbook = sheet.getWorkbook();
        CellStyle headerStyle = createHeaderStyle(workbook);
        CellStyle dataStyle = createDataStyle(workbook);
        CellStyle centerStyle = createCenterDataStyle(workbook);

        Row headerRow = sheet.getRow(0);
        if (headerRow != null) {
            for (int c = 0; c < 4; c++) {
                Cell cell = headerRow.getCell(c);
                if (cell != null) {
                    cell.setCellStyle(headerStyle);
                }
            }
        }

        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            applyStyle(row.getCell(0), centerStyle);
            applyStyle(row.getCell(1), dataStyle);
            applyStyle(row.getCell(2), centerStyle);
            applyStyle(row.getCell(3), centerStyle);
        }

        autoSizeColumns(sheet, 4);
        sheet.createFreezePane(0, 1);
        sheet.setAutoFilter(new CellRangeAddress(0, Math.max(sheet.getLastRowNum(), 0), 0, 3));
    }

    private static CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontName(FONT_NAME);
        font.setFontHeightInPoints(FONT_SIZE);
        font.setBold(true);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setWrapText(true);
        applyBorders(style);
        return style;
    }

    private static CellStyle createDataStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontName(FONT_NAME);
        font.setFontHeightInPoints(FONT_SIZE);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(false);
        applyBorders(style);
        return style;
    }

    private static CellStyle createCenterDataStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontName(FONT_NAME);
        font.setFontHeightInPoints(FONT_SIZE);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(false);
        applyBorders(style);
        return style;
    }

    private static void applyBorders(CellStyle style) {
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
    }

    private static void applyStyle(Cell cell, CellStyle style) {
        if (cell != null) {
            cell.setCellStyle(style);
        }
    }

    private static void autoSizeColumns(Sheet sheet, int columnCount) {
        for (int c = 0; c < columnCount; c++) {
            int maxLength = 0;
            for (int r = 0; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) {
                    continue;
                }
                Cell cell = row.getCell(c);
                if (cell == null) {
                    continue;
                }
                String value = getCellText(cell);
                maxLength = Math.max(maxLength, value.length());
            }
            int width = Math.max(MIN_COLUMN_WIDTH, Math.min((maxLength + 4) * 256, MAX_COLUMN_WIDTH));
            sheet.setColumnWidth(c, width);
        }
    }

    private static String getCellText(Cell cell) {
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf((long) cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            default -> cell.toString();
        };
    }
}
