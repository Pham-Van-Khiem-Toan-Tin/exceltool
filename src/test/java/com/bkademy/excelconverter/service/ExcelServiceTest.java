package com.bkademy.excelconverter.service;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.springframework.mock.web.MockMultipartFile;

import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.nio.file.Path;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import static org.junit.jupiter.api.Assertions.*;

class ExcelServiceTest {

    private static final Path DRL_SAMPLE = Path.of(
            "KỲ I (21-22)-20260619T031255Z-3-001",
            "KỲ I (21-22)",
            "ĐRL kì 1 năm 21-22.xlsx"
    );

    private static final Path RHM_SAMPLE = Path.of(
            "KỲ I (21-22)-20260619T031255Z-3-001",
            "KỲ I (21-22)",
            "Danh sách công nhận điểm rèn luyện RHM HKI_21.22.xlsx"
    );

    private static final Path Y6RHM_SAMPLE = Path.of(
            "drive-download-20260619T035347Z-3-001",
            "KỲ II (23-24)",
            "BẢNG THỐNG KÊ ĐIỂM RÈN LUYỆN KÌ II Y6RHM 2018 - 2024.xlsx"
    );

    @Test
    void extractsStudentDataFromDrlSampleFile() throws Exception {
        processAndAssertSample(DRL_SAMPLE, "20211", "215101_YHT0002");
    }

    @Test
    void extractsStudentDataFromRhmSampleFile() throws Exception {
        processAndAssertSample(RHM_SAMPLE, "20201", "2055010004");
    }

    @Test
    void extractsStudentDataFromY6RhmSampleFile() throws Exception {
        processAndAssertSample(Y6RHM_SAMPLE, "20232", "1855010004");
    }

    @Test
    void keepsEachInputSheetSeparateInOutput() throws Exception {
        ExcelService service = new ExcelService();

        try (Workbook input = new XSSFWorkbook();
             ByteArrayInputStream inputStream = createMultiSheetWorkbook(input)) {
            MockMultipartFile file = new MockMultipartFile(
                    "files",
                    "multi-sheet.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    inputStream.readAllBytes()
            );
            byte[] zipBytes = service.processFiles(new MockMultipartFile[]{file});

            try (ZipInputStream zis = new ZipInputStream(new ByteArrayInputStream(zipBytes))) {
                assertNotNull(zis.getNextEntry());

                Workbook result = WorkbookFactory.create(zis);
                assertEquals(2, result.getNumberOfSheets());
                assertNotNull(result.getSheet("Lop A"));
                assertNotNull(result.getSheet("Lop B"));
                assertEquals(1, (int) result.getSheet("Lop A").getRow(1).getCell(0).getNumericCellValue());
                assertEquals(1, (int) result.getSheet("Lop B").getRow(1).getCell(0).getNumericCellValue());
                assertEquals("SV001", result.getSheet("Lop A").getRow(1).getCell(1).getStringCellValue());
                assertEquals("SV101", result.getSheet("Lop B").getRow(1).getCell(1).getStringCellValue());
            }
        }
    }

    private ByteArrayInputStream createMultiSheetWorkbook(Workbook workbook) throws Exception {
        createStudentSheet(workbook, "Lop A", "SV001", "85");
        createStudentSheet(workbook, "Lop B", "SV101", "90");

        try (var out = new java.io.ByteArrayOutputStream()) {
            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        }
    }

    private void createStudentSheet(Workbook workbook, String sheetName, String mssv, String diem) {
        Sheet sheet = workbook.createSheet(sheetName);
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("MSSV");
        header.createCell(1).setCellValue("Điểm rèn luyện");

        Row student = sheet.createRow(1);
        student.createCell(0).setCellValue(mssv);
        student.createCell(1).setCellValue(diem);
    }

    private void processAndAssertSample(Path sampleFile, String expectedSemester, String expectedMssvPrefix) throws Exception {
        ExcelService service = new ExcelService();
        byte[] zipBytes;

        try (FileInputStream fis = new FileInputStream(sampleFile.toFile())) {
            MockMultipartFile file = new MockMultipartFile(
                    "files",
                    sampleFile.getFileName().toString(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    fis
            );
            zipBytes = service.processFiles(new MockMultipartFile[]{file});
        }

        try (ZipInputStream zis = new ZipInputStream(new ByteArrayInputStream(zipBytes))) {
            ZipEntry entry = zis.getNextEntry();
            assertNotNull(entry);

            Workbook result = WorkbookFactory.create(zis);
            assertTrue(result.getNumberOfSheets() >= 1, "Expected at least one output sheet");
            Sheet sheet = result.getSheetAt(0);
            assertNotNull(sheet);
            assertTrue(sheet.getLastRowNum() > 1, "Expected student rows in output");

            Row firstStudent = sheet.getRow(1);
            assertEquals(1, (int) firstStudent.getCell(0).getNumericCellValue());
            assertTrue(firstStudent.getCell(1).getStringCellValue().startsWith(expectedMssvPrefix.substring(0, 4)));
            assertFalse(firstStudent.getCell(2).getStringCellValue().isBlank());
            assertEquals(expectedSemester, firstStudent.getCell(3).getStringCellValue());

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                String mssv = sheet.getRow(r).getCell(1).getStringCellValue();
                assertFalse(mssv.toLowerCase().contains("học kỳ"), "Title row must not be treated as MSSV: " + mssv);
                assertFalse(mssv.toLowerCase().contains("quyết định"), "Decision row must not be treated as MSSV: " + mssv);
            }
        }
    }
}
