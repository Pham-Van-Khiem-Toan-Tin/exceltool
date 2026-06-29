package com.bkademy.excelconverter.service;

import org.apache.poi.ss.formula.BaseFormulaEvaluator;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.nio.file.*;
import java.text.Normalizer;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

@Service
public class ExcelService {

    private record Student(String mssv, String diem, String kyHoc, String sourceFile, String sheetName) {
        public String getConflictKey() {
            return mssv + "::" + kyHoc;
        }
    }

    public byte[] processFiles(MultipartFile[] files) throws IOException {
        List<Student> allStudents = new ArrayList<>();
        Map<String, byte[]> originalFileContents = new HashMap<>();

        for (MultipartFile file : files) {
            String filename = file.getOriginalFilename();
            if (filename == null) continue;

            byte[] content = file.getBytes();
            originalFileContents.put(filename, content);

            if (filename.toLowerCase().endsWith(".zip")) {
                extractStudentsFromZip(new ByteArrayInputStream(content), filename, allStudents);
            } else if (isExcelFile(filename)) {
                extractStudentsFromExcel(new ByteArrayInputStream(content), filename, allStudents);
            }
        }

        return generateZipOutput(allStudents, originalFileContents);
    }

    public byte[] processDirectory(Path rootDir) throws IOException {
        List<Student> allStudents = new ArrayList<>();
        Map<String, byte[]> originalFileContents = new HashMap<>();
        Path baseDir = rootDir.toAbsolutePath().normalize();

        try (Stream<Path> paths = Files.walk(baseDir)) {
            List<Path> filePaths = paths.filter(Files::isRegularFile)
                    .filter(this::isSupportedFile)
                    .sorted()
                    .collect(Collectors.toList());

            for (Path path : filePaths) {
                String relativePath = baseDir.relativize(path).toString().replace('\\', '/');
                byte[] content = Files.readAllBytes(path);
                originalFileContents.put(relativePath, content);

                if (relativePath.toLowerCase().endsWith(".zip")) {
                    extractStudentsFromZip(new ByteArrayInputStream(content), relativePath, allStudents);
                } else if (isExcelFile(relativePath)) {
                    extractStudentsFromExcel(new ByteArrayInputStream(content), relativePath, allStudents);
                }
            }
        }

        return generateZipOutput(allStudents, originalFileContents);
    }

    private void extractStudentsFromExcel(InputStream excelInput, String filePath, List<Student> allStudents) throws IOException {
        try (Workbook workbook = WorkbookFactory.create(excelInput)) {
            DataFormatter formatter = new DataFormatter();
            FormulaEvaluator evaluator = createFormulaEvaluator(workbook);
            String fileName = stripExtension(Paths.get(filePath).getFileName().toString());
            String folderName = Optional.ofNullable(Paths.get(filePath).getParent()).map(p -> p.getFileName().toString()).orElse("");

            Set<String> usedSheetNames = new HashSet<>();

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                String originalSheetName = sheet.getSheetName();
                String uniqueSheetName = uniqueSheetName(originalSheetName, usedSheetNames);
                allStudents.addAll(extractStudentsFromSheet(sheet, formatter, evaluator, fileName, folderName, filePath, uniqueSheetName));
            }
        } catch (Exception e) {
            System.err.println("Could not process Excel file: " + filePath + ". Error: " + e.getMessage());
        }
    }

    private void extractStudentsFromZip(InputStream zipInput, String zipFilePath, List<Student> allStudents) throws IOException {
        try (ZipInputStream zis = new ZipInputStream(zipInput)) {
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                if (!entry.isDirectory() && isExcelFile(entry.getName())) {
                    ByteArrayOutputStream entryOut = new ByteArrayOutputStream();
                    zis.transferTo(entryOut);
                    String entryPath = zipFilePath + "/" + entry.getName();
                    extractStudentsFromExcel(new ByteArrayInputStream(entryOut.toByteArray()), entryPath, allStudents);
                }
                zis.closeEntry();
            }
        }
    }

    private byte[] generateZipOutput(List<Student> allStudents, Map<String, byte[]> originalFileContents) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        try (ZipOutputStream zos = new ZipOutputStream(baos)) {

            Map<String, Set<String>> scoresByStudentAndSemester = new HashMap<>();
            for (Student s : allStudents) {
                scoresByStudentAndSemester.computeIfAbsent(s.getConflictKey(), k -> new HashSet<>()).add(s.diem());
            }

            Set<String> conflictKeys = scoresByStudentAndSemester.entrySet().stream()
                    .filter(e -> e.getValue().size() > 1)
                    .map(Map.Entry::getKey)
                    .collect(Collectors.toSet());

            List<Student> conflictingStudents = allStudents.stream()
                    .filter(s -> conflictKeys.contains(s.getConflictKey()))
                    .sorted(Comparator.comparing(Student::mssv)
                            .thenComparing(Student::kyHoc)
                            .thenComparing(Student::diem))
                    .collect(Collectors.toList());

            List<Student> cleanStudents = allStudents.stream()
                    .filter(s -> !conflictKeys.contains(s.getConflictKey()))
                    .collect(Collectors.toList());

            Map<String, List<Student>> cleanStudentsByFile = cleanStudents.stream()
                    .collect(Collectors.groupingBy(Student::sourceFile));

            for (Map.Entry<String, byte[]> fileEntry : originalFileContents.entrySet()) {
                String filePath = fileEntry.getKey();
                if (!isExcelFile(filePath)) {
                    zos.putNextEntry(new ZipEntry(filePath));
zos.write(fileEntry.getValue());
                    zos.closeEntry();
                }
            }

            for (Map.Entry<String, List<Student>> entry : cleanStudentsByFile.entrySet()) {
                String sourceFile = entry.getKey();
                List<Student> fileStudents = entry.getValue();

                try (Workbook outputWorkbook = new XSSFWorkbook()) {
                    Map<String, List<Student>> studentsBySheet = fileStudents.stream()
                            .collect(Collectors.groupingBy(Student::sheetName));

                    for (Map.Entry<String, List<Student>> sheetEntry : studentsBySheet.entrySet()) {
                        Sheet outputSheet = outputWorkbook.createSheet(sheetEntry.getKey());
                        createOutputHeader(outputSheet, false);
                        int rowIdx = 1;
                        for (Student student : sheetEntry.getValue()) {
                            Row outRow = outputSheet.createRow(rowIdx++);
                            outRow.createCell(0).setCellValue(rowIdx - 1);
                            outRow.createCell(1).setCellValue(student.mssv());
                            outRow.createCell(2).setCellValue(student.diem());
                            outRow.createCell(3).setCellValue(student.kyHoc());
                        }
                        ExcelOutputFormatter.format(outputSheet);
                    }

                    if (outputWorkbook.getNumberOfSheets() > 0) {
                        ByteArrayOutputStream processedOut = new ByteArrayOutputStream();
                        outputWorkbook.write(processedOut);
                        String outputFilename = getOutputExcelFileName(sourceFile);
                        zos.putNextEntry(new ZipEntry(outputFilename));
                        zos.write(processedOut.toByteArray());
                        zos.closeEntry();
                    }
                }
            }

            if (!conflictingStudents.isEmpty()) {
                try (Workbook conflictsWorkbook = new XSSFWorkbook()) {
                    Sheet conflictsSheet = conflictsWorkbook.createSheet("Sinh viên khác điểm");
                    createOutputHeader(conflictsSheet, true);
                    int rowIdx = 1;
                    for (Student student : conflictingStudents) {
                        Row outRow = conflictsSheet.createRow(rowIdx++);
                        outRow.createCell(0).setCellValue(rowIdx - 1);
                        outRow.createCell(1).setCellValue(student.mssv());
                        outRow.createCell(2).setCellValue(student.diem());
                        outRow.createCell(3).setCellValue(student.kyHoc());
                        outRow.createCell(4).setCellValue(student.sourceFile());
                        outRow.createCell(5).setCellValue(student.sheetName());
                    }
                    ExcelOutputFormatter.format(conflictsSheet);

                    ByteArrayOutputStream conflictsOut = new ByteArrayOutputStream();
                    conflictsWorkbook.write(conflictsOut);
                    zos.putNextEntry(new ZipEntry("[MAU-THUAN] Sinh vien khac diem.xlsx"));
                    zos.write(conflictsOut.toByteArray());
                    zos.closeEntry();
                }
            }
        }
        return baos.toByteArray();
    }

    private String getOutputExcelFileName(String inputPath) {
        String filename = Paths.get(inputPath).getFileName().toString();
        String stripped = stripExtension(filename);
        String parent = Optional.ofNullable(Paths.get(inputPath).getParent()).map(Path::toString).orElse("");
        parent = parent.replace('\\', '/');
        if (!parent.isEmpty() && !parent.endsWith("/")) {
            parent += "/";
        }
        return parent + stripped + ".xlsx";
    }

    private boolean isExcelFile(String filename) {
        String lower = filename.toLowerCase();
        return lower.endsWith(".xlsx") || lower.endsWith(".xls");
    }

    private boolean isSupportedFile(Path path) {
        String name = path.getFileName().toString().toLowerCase();
        return name.endsWith(".xlsx") || name.endsWith(".xls") || name.endsWith(".zip");
    }

    private void createOutputHeader(Sheet outputSheet, boolean isConflictSheet) {
        Row outHeader = outputSheet.createRow(0);
        outHeader.createCell(0).setCellValue("STT");
        outHeader.createCell(1).setCellValue("MSSV");
        outHeader.createCell(2).setCellValue("Điểm rèn luyện");
        outHeader.createCell(3).setCellValue("Kỳ học");
        if (isConflictSheet) {
            outHeader.createCell(4).setCellValue("File gốc");
            outHeader.createCell(5).setCellValue("Sheet gốc");
        }
    }

    private String uniqueSheetName(String rawName, Set<String> usedNames) {
        String sanitized = sanitizeSheetName(rawName);
        if (sanitized.isEmpty()) {
            sanitized = "Sheet";
        }
        String candidate = sanitized;
        int suffix = 2;
        while (usedNames.contains(candidate)) {
            String suffixText = " (" + suffix + ")";
            int maxBaseLength = 31 - suffixText.length();
            String base = sanitized.length() > maxBaseLength ? sanitized.substring(0, maxBaseLength) : sanitized;
            candidate = base + suffixText;
            suffix++;
        }
        usedNames.add(candidate);
        return candidate;
    }

    private String sanitizeSheetName(String name) {
        if (name == null || name.isBlank()) {
            return "";
        }
        String sanitized = name.replaceAll("[\\\\/?*\\[\\]]", " ").trim();
        return sanitized.length() > 31 ? sanitized.substring(0, 31) : sanitized;
    }

    private List<Student> extractStudentsFromSheet(Sheet inputSheet, DataFormatter formatter,
                                                   FormulaEvaluator evaluator, String fileName, String folderName, String sourceFile, String sheetName) {
        int headerRowIdx = -1;
        int mssvCol = -1;
        int diemCol = -1;
        int kyHocCol = -1;

        int lastRow = Math.min(50, inputSheet.getLastRowNum());
        for (int r = 0; r <= lastRow; r++) {
            Row row = inputSheet.getRow(r);
            if (row == null) continue;

            int tempMssv = -1, tempDiem = -1, tempKyHoc = -1;
            int lastCol = row.getLastCellNum();
            for (int c = 0; c < lastCol; c++) {
                String val = getCellValueAsString(row.getCell(c), formatter, evaluator);
                if (isMssvColumn(val)) tempMssv = c;
                if (isDiemRenLuyenColumn(val)) tempDiem = c;
                if (isKyHocColumn(val)) tempKyHoc = c;
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
            return Collections.emptyList();
        }

        List<Student> students = new ArrayList<>();
        String titleSemesterText = findSemesterTitle(inputSheet, headerRowIdx, formatter, evaluator);

        for (int r = headerRowIdx + 1; r <= inputSheet.getLastRowNum(); r++) {
            Row row = inputSheet.getRow(r);
            if (row == null) continue;

            String mssv = getCellValueAsString(row.getCell(mssvCol), formatter, evaluator).trim();
            if (!isValidStudentRow(mssv)) continue;

            String diem = getCellValueAsString(row.getCell(diemCol), formatter, evaluator).trim();
            String columnKyHoc = kyHocCol != -1 ? getCellValueAsString(row.getCell(kyHocCol), formatter, evaluator).trim() : "";
            String kyHoc = SemesterParser.resolve(columnKyHoc, titleSemesterText, fileName, folderName);

            students.add(new Student(mssv, diem, kyHoc, sourceFile, sheetName));
        }

        return students;
    }

    private String findSemesterTitle(Sheet sheet, int headerRowIdx, DataFormatter formatter, FormulaEvaluator evaluator) {
        for (int r = 0; r < headerRowIdx; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            for (int c = 0; c < row.getLastCellNum(); c++) {
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
        if (mssv.isEmpty()) return false;
        String n = normalize(mssv);
        if (n.contains("tong") || n.contains("hoc ky") || n.contains("hoc ki")
                || n.contains("quyet dinh") || n.contains("kem theo") || n.contains("nam hoc")) {
            return false;
        }
        return mssv.matches("[\\d\\w._-]{4,25}");
    }

    private boolean isMssvColumn(String val) {
        String n = normalize(val);
        if (n.length() > 40) return false;
        return n.equals("mssv") || n.equals("msv")
                || n.contains("ma so")
                || (n.matches(".*\\bma\\b.*") && (n.contains("sinh vien") || n.endsWith(" sv")));
    }

    private boolean isDiemRenLuyenColumn(String val) {
        String n = normalize(val);
        if (n.length() > 40) return false;
        return n.equals("drl") || n.equals("diem") || n.equals("diem rl")
                || (n.contains("diem") && n.contains("ren luyen"))
                || (n.contains("diem") && n.contains("rl"));
    }

    private boolean isKyHocColumn(String val) {
        String n = normalize(val);
        if (n.length() > 40) return false;
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
        if (cell == null) return "";
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