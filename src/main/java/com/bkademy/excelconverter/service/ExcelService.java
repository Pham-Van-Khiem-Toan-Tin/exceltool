package com.bkademy.excelconverter.service;

import org.apache.poi.ss.formula.BaseFormulaEvaluator;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.nio.file.*;
import java.text.Normalizer;
import java.util.stream.Stream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

@Service
public class ExcelService {

    public byte[] processFiles(MultipartFile[] files) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();

        try (ZipOutputStream zos = new ZipOutputStream(baos)) {
            for (MultipartFile file : files) {
                String filename = file.getOriginalFilename();

                if (filename == null) continue;

                if (filename.toLowerCase().endsWith(".zip")) {
                    processZipFile(file.getInputStream(), zos);
                } else if (filename.toLowerCase().endsWith(".xls") || filename.toLowerCase().endsWith(".xlsx")) {
                    processAndAddExcelToZip(file.getInputStream(), filename, zos);
                }
            }
        }

        return baos.toByteArray();
    }

    public byte[] processDirectory(Path rootDir) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        Path baseDir = rootDir.toAbsolutePath().normalize();

        try (ZipOutputStream zos = new ZipOutputStream(baos);
             Stream<Path> paths = Files.walk(baseDir)) {
            paths.filter(Files::isRegularFile)
                    .filter(this::isSupportedFile)
                    .sorted()
                    .forEach(path -> {
                        String relativePath = baseDir.relativize(path).toString().replace('\\', '/');
                        try (InputStream in = Files.newInputStream(path)) {
                            if (relativePath.toLowerCase().endsWith(".zip")) {
                                processZipFile(in, zos);
                            } else {
                                processAndAddExcelToZip(in, relativePath, zos);
                            }
                        } catch (IOException e) {
                            throw new UncheckedIOException(e);
                        }
                    });
        } catch (UncheckedIOException e) {
            throw e.getCause();
        }

        return baos.toByteArray();
    }

    private boolean isSupportedFile(Path path) {
        String name = path.getFileName().toString().toLowerCase();
        return name.endsWith(".xlsx") || name.endsWith(".xls") || name.endsWith(".zip");
    }

    private void processZipFile(InputStream zipInput, ZipOutputStream zos) throws IOException {
        try (ZipInputStream zis = new ZipInputStream(zipInput)) {
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                String entryName = entry.getName();

                if (entry.isDirectory()) {
                    zos.putNextEntry(new ZipEntry(entryName));
                    zos.closeEntry();
                } else {
                    if (entryName.toLowerCase().endsWith(".xls") || entryName.toLowerCase().endsWith(".xlsx")) {
                        ByteArrayOutputStream entryOut = new ByteArrayOutputStream();
                        zis.transferTo(entryOut);
                        ByteArrayInputStream entryIn = new ByteArrayInputStream(entryOut.toByteArray());

                        processAndAddExcelToZip(entryIn, entryName, zos);
                    } else {
                        zos.putNextEntry(new ZipEntry(entryName));
                        zis.transferTo(zos);
                        zos.closeEntry();
                    }
                }
                zis.closeEntry();
            }
        }
    }

    private void processAndAddExcelToZip(InputStream excelInput, String filePath, ZipOutputStream zos) throws IOException {
        try (Workbook inputWorkbook = WorkbookFactory.create(excelInput);
             Workbook outputWorkbook = extractData(inputWorkbook, filePath);
             ByteArrayOutputStream processedOut = new ByteArrayOutputStream()) {

            outputWorkbook.write(processedOut);

            String filename = Paths.get(filePath).getFileName().toString();
            if (filename.toLowerCase().endsWith(".xls")) {
                filename = filename.substring(0, filename.length() - 4) + ".xlsx";
            }

            String parentPath = Paths.get(filePath).getParent() != null
                    ? Paths.get(filePath).getParent().toString().replace('\\', '/')
                    : "";
            String zipEntryName = parentPath.isEmpty() ? filename : parentPath + "/" + filename;

            ZipEntry zipEntry = new ZipEntry(zipEntryName);
            zos.putNextEntry(zipEntry);
            zos.write(processedOut.toByteArray());
            zos.closeEntry();
        }
    }

    private Workbook extractData(Workbook inputWorkbook, String filePath) {
        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet = outputWorkbook.createSheet("Kết quả");

        Row outHeader = outputSheet.createRow(0);
        outHeader.createCell(0).setCellValue("STT");
        outHeader.createCell(1).setCellValue("MSSV");
        outHeader.createCell(2).setCellValue("Điểm rèn luyện");
        outHeader.createCell(3).setCellValue("Kỳ học");

        String fileName = stripExtension(Paths.get(filePath).getFileName().toString());
        String folderName = Paths.get(filePath).getParent() != null
                ? Paths.get(filePath).getParent().getFileName().toString()
                : "";

        DataFormatter formatter = new DataFormatter();
        FormulaEvaluator evaluator = createFormulaEvaluator(inputWorkbook);
        int outRowIdx = 1;

        for (int i = 0; i < inputWorkbook.getNumberOfSheets(); i++) {
            Sheet inputSheet = inputWorkbook.getSheetAt(i);
            outRowIdx = extractSheet(inputSheet, outputSheet, formatter, evaluator, outRowIdx, fileName, folderName);
        }

        ExcelOutputFormatter.format(outputSheet);
        return outputWorkbook;
    }

    private int extractSheet(Sheet inputSheet, Sheet outputSheet, DataFormatter formatter,
                             FormulaEvaluator evaluator, int outRowIdx, String fileName, String folderName) {
        int headerRowIdx = -1;
        int mssvCol = -1;
        int diemCol = -1;
        int kyHocCol = -1;

        int lastRow = Math.min(50, inputSheet.getLastRowNum());
        for (int r = 0; r <= lastRow; r++) {
            Row row = inputSheet.getRow(r);
            if (row == null) continue;

            int tempMssv = -1, tempDiem = -1, tempKyHoc = -1;
            int lastCol = Math.max(row.getLastCellNum(), 0);
            for (int c = 0; c < lastCol; c++) {
                String val = getCellValueAsString(row.getCell(c), formatter, evaluator);
                if (isMssvColumn(val)) {
                    tempMssv = c;
                }
                if (isDiemRenLuyenColumn(val)) {
                    tempDiem = c;
                }
                if (isKyHocColumn(val)) {
                    tempKyHoc = c;
                }
            }

            if (tempMssv != -1 && tempDiem != -1 && tempMssv != tempDiem) {
                headerRowIdx = r;
                mssvCol = tempMssv;
                diemCol = tempDiem;
                kyHocCol = tempKyHoc;
                break;
            }
        }

        if (headerRowIdx == -1) {
            return outRowIdx;
        }

        String titleSemesterText = findSemesterTitle(inputSheet, headerRowIdx, formatter, evaluator);

        for (int r = headerRowIdx + 1; r <= inputSheet.getLastRowNum(); r++) {
            Row row = inputSheet.getRow(r);
            if (row == null) continue;

            String mssv = getCellValueAsString(row.getCell(mssvCol), formatter, evaluator).trim();
            if (!isValidStudentRow(mssv)) continue;

            String diem = getCellValueAsString(row.getCell(diemCol), formatter, evaluator).trim();
            String columnKyHoc = kyHocCol != -1
                    ? getCellValueAsString(row.getCell(kyHocCol), formatter, evaluator).trim()
                    : "";
            String kyHoc = SemesterParser.resolve(columnKyHoc, titleSemesterText, fileName, folderName);

            Row outRow = outputSheet.createRow(outRowIdx);
            outRow.createCell(0).setCellValue(outRowIdx);
            outRow.createCell(1).setCellValue(mssv);
            outRow.createCell(2).setCellValue(diem);
            outRow.createCell(3).setCellValue(kyHoc);
            outRowIdx++;
        }

        return outRowIdx;
    }

    private String findSemesterTitle(Sheet sheet, int headerRowIdx, DataFormatter formatter, FormulaEvaluator evaluator) {
        for (int r = 0; r < headerRowIdx; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            int lastCol = Math.max(row.getLastCellNum(), 0);
            for (int c = 0; c < lastCol; c++) {
                String val = getCellValueAsString(row.getCell(c), formatter, evaluator).trim();
                String n = normalize(val);
                if (n.contains("hoc ky") || n.contains("hoc ki") || n.contains("nam hoc")) {
                    return val;
                }
            }
        }
        return "";
    }

    private String stripExtension(String filename) {
        int dot = filename.lastIndexOf('.');
        return dot > 0 ? filename.substring(0, dot) : filename;
    }

    private boolean isValidStudentRow(String mssv) {
        if (mssv.isEmpty()) {
            return false;
        }
        String n = normalize(mssv);
        if (n.contains("tong") || n.contains("hoc ky") || n.contains("hoc ki")
                || n.contains("quyet dinh") || n.contains("kem theo") || n.contains("nam hoc")) {
            return false;
        }
        return mssv.matches("[\\d\\w._-]{4,25}");
    }

    private boolean isMssvColumn(String val) {
        String n = normalize(val);
        if (n.length() > 40) {
            return false;
        }
        return n.equals("mssv") || n.equals("msv")
                || n.contains("ma so")
                || (n.matches(".*\\bma\\b.*") && (n.contains("sinh vien") || n.endsWith(" sv")));
    }

    private boolean isDiemRenLuyenColumn(String val) {
        String n = normalize(val);
        if (n.length() > 40) {
            return false;
        }
        return n.equals("drl") || n.equals("diem") || n.equals("diem rl")
                || (n.contains("diem") && n.contains("ren luyen"))
                || (n.contains("diem") && n.contains("rl"));
    }

    private boolean isKyHocColumn(String val) {
        String n = normalize(val);
        if (n.length() > 40) {
            return false;
        }
        return n.contains("ky hoc") || n.contains("hoc ky") || n.contains("hoc ki");
    }

    private String normalize(String text) {
        if (text == null || text.isEmpty()) return "";
        String n = Normalizer.normalize(text, Normalizer.Form.NFD);
        n = n.replaceAll("\\p{InCombiningDiacriticalMarks}+", "");
        n = n.replace('đ', 'd').replace('Đ', 'd');
        return n.toLowerCase().trim().replaceAll("\\s+", " ");
    }

    private FormulaEvaluator createFormulaEvaluator(Workbook workbook) {
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        if (evaluator instanceof BaseFormulaEvaluator baseEvaluator) {
            baseEvaluator.setIgnoreMissingWorkbooks(true);
        }
        return evaluator;
    }

    private String getCellValueAsString(Cell cell, DataFormatter formatter, FormulaEvaluator evaluator) {
        if (cell == null) {
            return "";
        }
        try {
            return formatter.formatCellValue(cell, evaluator).trim();
        } catch (Exception ex) {
            return readCellWithoutExternalLinks(cell, formatter);
        }
    }

    private String readCellWithoutExternalLinks(Cell cell, DataFormatter formatter) {
        if (cell.getCellType() == CellType.FORMULA) {
            return switch (cell.getCachedFormulaResultType()) {
                case STRING -> cell.getRichStringCellValue().getString().trim();
                case NUMERIC -> formatNumeric(cell.getNumericCellValue());
                case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
                default -> formatter.formatCellValue(cell).trim();
            };
        }
        return formatter.formatCellValue(cell).trim();
    }

    private String formatNumeric(double value) {
        if (value == Math.floor(value)) {
            return String.valueOf((long) value);
        }
        return String.valueOf(value);
    }
}
