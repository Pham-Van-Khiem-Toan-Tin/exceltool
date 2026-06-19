package com.bkademy.excelconverter.service;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.*;

class SemesterParserTest {

    @ParameterizedTest
    @CsvSource({
            "HOC KY I NAM HOC 2020 - 2021, 20201",
            "Học kỳ I năm học 2020-2021, 20201",
            "HOC KY II NAM HOC 2020-2021, 20202",
            "Học kì 2 năm học 2020 - 2021, 20202",
            "HOC KY I NAM HOC 2021 - 2022, 20211",
            "ĐRL kì 1 năm 21-22, 20211",
            "KỲ I (21-22), 20211",
            "KHÓA 2020.KỲ I.21-22, 20211",
            "HK1-2021-2022, 20211",
            "20201, 20201",
            "HOC KY II NAM HOC 2023 - 2024, 20232",
            "BANG THONG KE DIEM REN LUYEN KI II Y6RHM 2018 - 2024, ''"
    })
    void parseSemesterCode(String input, String expected) {
        assertEquals(expected.isEmpty() ? "" : expected, SemesterParser.parse(input));
    }

    @Test
    void usesColumnValueWhenPresent() {
        assertEquals("20299", SemesterParser.resolve("20299", "Học kỳ I năm học 2020-2021", "file", "folder"));
    }

    @Test
    void fallsBackToTitleThenFileThenFolder() {
        assertEquals("20201", SemesterParser.resolve("", "Học kỳ I năm học 2020-2021", "unknown", "unknown"));
        assertEquals("20211", SemesterParser.resolve("", "", "ĐRL kì 1 năm 21-22", "unknown"));
        assertEquals("20211", SemesterParser.resolve("", "", "unknown", "KỲ I (21-22)"));
        assertEquals("", SemesterParser.resolve("", "", "unknown", "unknown"));
    }
}
