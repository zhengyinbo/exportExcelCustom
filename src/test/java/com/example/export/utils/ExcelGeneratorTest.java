package com.example.export.utils;

import com.example.export.model.SheetData;
import com.example.export.model.UserVO;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

public class ExcelGeneratorTest {

    @Test
    public void testGenerateMultiSheet() throws IllegalAccessException {
        List<UserVO> group1 = new ArrayList<>();
        group1.add(
                new UserVO(1L, "Alice", 30, "alice@example.com", "1234567890", LocalDate.now(), LocalDateTime.now()));

        List<UserVO> group2 = new ArrayList<>();
        group2.add(new UserVO(2L, "Bob", 25, "bob@example.com", "0987654321", LocalDate.now(), LocalDateTime.now()));

        List<SheetData> sheets = new ArrayList<>();
        sheets.add(new SheetData("Group 1", group1, UserVO.class));
        sheets.add(new SheetData("Group 2", group2, UserVO.class));

        Workbook workbook = ExcelGenerator.generateMultiSheet(sheets);

        assertNotNull(workbook);
        assertEquals(2, workbook.getNumberOfSheets());
        assertEquals("Group 1", workbook.getSheetName(0));
        assertEquals("Group 2", workbook.getSheetName(1));
        assertEquals(1, workbook.getSheetAt(0).getLastRowNum() - 1); // 2 header rows + 1 data row - 1 (0-indexed) = 2.
                                                                     // Wait. 0,1 are headers. 2 is data. LastRowNum is
                                                                     // 2.
        // Row 0: Header 1
        // Row 1: Header 2
        // Row 2: Data 1
        // LastRowNum should be 2.
        assertEquals(2, workbook.getSheetAt(0).getLastRowNum());
    }
}
