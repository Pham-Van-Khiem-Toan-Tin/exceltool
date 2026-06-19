package com.bkademy.excelconverter.service;

import java.text.Normalizer;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

final class SemesterParser {

    private static final Pattern EXISTING_CODE = Pattern.compile("^20\\d{3}$");
    private static final Pattern FULL_YEAR_RANGE = Pattern.compile("(20\\d{2})\\s*[-–/]\\s*(20\\d{2})");
    private static final Pattern SHORT_YEAR_RANGE = Pattern.compile("\\b(\\d{2})\\s*[-–/]\\s*(\\d{2})\\b");
    private static final Pattern SEMESTER_ROMAN = Pattern.compile("\\b(?:hoc\\s+)?(?:ky|ki)\\s*(i{1,2})\\b");
    private static final Pattern SEMESTER_NUMBER = Pattern.compile("\\b(?:hoc\\s+)?(?:ky|ki|hk)\\s*([12])\\b");
    private static final Pattern HK_PREFIX = Pattern.compile("\\bhk\\s*([12])\\b");

    private SemesterParser() {
    }

    static String resolve(String columnValue, String titleText, String fileName, String folderName) {
        if (columnValue != null && !columnValue.isBlank()) {
            return columnValue.trim();
        }

        for (String source : new String[]{titleText, fileName, folderName}) {
            String code = parse(source);
            if (!code.isEmpty()) {
                return code;
            }
        }
        return "";
    }

    static String parse(String text) {
        if (text == null || text.isBlank()) {
            return "";
        }

        String trimmed = text.trim();
        if (EXISTING_CODE.matcher(trimmed).matches()) {
            return trimmed;
        }

        String normalized = normalize(trimmed);
        Integer startYear = extractStartYear(normalized);
        Integer semester = extractSemester(normalized);

        if (startYear != null && semester != null) {
            return startYear + String.valueOf(semester);
        }
        return "";
    }

    private static Integer extractStartYear(String normalized) {
        Matcher full = FULL_YEAR_RANGE.matcher(normalized);
        while (full.find()) {
            int start = Integer.parseInt(full.group(1));
            int end = Integer.parseInt(full.group(2));
            if (end == start + 1) {
                return start;
            }
        }

        Matcher shortRange = SHORT_YEAR_RANGE.matcher(normalized);
        while (shortRange.find()) {
            int first = Integer.parseInt(shortRange.group(1));
            int second = Integer.parseInt(shortRange.group(2));
            if (second == first + 1 || (first == 99 && second == 0)) {
                return 2000 + first;
            }
        }
        return null;
    }

    private static Integer extractSemester(String normalized) {
        Matcher roman = SEMESTER_ROMAN.matcher(normalized);
        if (roman.find()) {
            return roman.group(1).length() == 1 ? 1 : 2;
        }

        Matcher hk = HK_PREFIX.matcher(normalized);
        if (hk.find()) {
            return Integer.parseInt(hk.group(1));
        }

        Matcher number = SEMESTER_NUMBER.matcher(normalized);
        if (number.find()) {
            return Integer.parseInt(number.group(1));
        }

        if (normalized.contains("ky i ") || normalized.endsWith("ky i")
                || normalized.contains("ki i ") || normalized.endsWith("ki i")) {
            return 1;
        }
        if (normalized.contains("ky ii") || normalized.contains("ki ii")) {
            return 2;
        }

        return null;
    }

    private static String normalize(String text) {
        String n = Normalizer.normalize(text, Normalizer.Form.NFD);
        n = n.replaceAll("\\p{InCombiningDiacriticalMarks}+", "");
        n = n.replace('đ', 'd').replace('Đ', 'd');
        return n.toLowerCase().trim().replaceAll("\\s+", " ");
    }
}
