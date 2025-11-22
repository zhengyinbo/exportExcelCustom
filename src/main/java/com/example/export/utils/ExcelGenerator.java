package com.example.export.utils;

import com.example.export.annotation.ExcelField;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

public class ExcelGenerator {

    public static Workbook generate(List<?> data, Class<?> clazz, String sheetName, Set<String> exportFields)
            throws IllegalAccessException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(sheetName);

        // Get fields annotated with @ExcelField
        List<Field> fields = Arrays.stream(clazz.getDeclaredFields())
                .filter(f -> f.isAnnotationPresent(ExcelField.class))
                .filter(f -> {
                    if (exportFields == null || exportFields.isEmpty()) {
                        return true;
                    }
                    // Filter by field name or title? Requirement says "export fields", usually
                    // implies field name.
                    // Let's support matching field name.
                    return exportFields.contains(f.getName());
                })
                .sorted(Comparator.comparingInt(f -> f.getAnnotation(ExcelField.class).order()))
                .collect(Collectors.toList());

        if (fields.isEmpty()) {
            return workbook;
        }

        // Create styles
        CellStyle headerStyle = createHeaderStyle(workbook);
        CellStyle dataStyle = createDataStyle(workbook);

        // Create Headers
        createHeaders(sheet, fields, headerStyle);

        // Fill Data
        fillData(sheet, fields, data, dataStyle);

        // Auto-size columns (optional, can be slow for large data)
        // for (int i = 0; i < fields.size(); i++) {
        // sheet.autoSizeColumn(i);
        // }

        // Set manual width
        for (int i = 0; i < fields.size(); i++) {
            ExcelField excelField = fields.get(i).getAnnotation(ExcelField.class);
            sheet.setColumnWidth(i, excelField.width() * 256);
        }

        return workbook;
    }

    private static void createHeaders(Sheet sheet, List<Field> fields, CellStyle headerStyle) {
        Row row1 = sheet.createRow(0);
        Row row2 = sheet.createRow(1);

        boolean hasGroup = fields.stream().anyMatch(f -> !f.getAnnotation(ExcelField.class).group().isEmpty());

        // If no groups at all, we can just have one header row, but requirement says 2
        // levels.
        // We will always create 2 rows. If no group, merge vertically.

        String currentGroup = null;
        int groupStartCol = 0;

        for (int i = 0; i < fields.size(); i++) {
            Field field = fields.get(i);
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            String group = excelField.group();
            String title = excelField.title();

            // L1 Header
            Cell cell1 = row1.createCell(i);
            cell1.setCellStyle(headerStyle);

            // L2 Header
            Cell cell2 = row2.createCell(i);
            cell2.setCellValue(title);
            cell2.setCellStyle(headerStyle);

            // Handle Grouping logic
            if (group.isEmpty()) {
                // No group, merge vertically (Row 0 and Row 1)
                cell1.setCellValue(title); // Put title in top cell too for visual if merge fails (but we merge)
                sheet.addMergedRegion(new CellRangeAddress(0, 1, i, i));

                // If we were tracking a group, close it
                if (currentGroup != null) {
                    if (i - 1 > groupStartCol) {
                        sheet.addMergedRegion(new CellRangeAddress(0, 0, groupStartCol, i - 1));
                    }
                    currentGroup = null;
                }
            } else {
                cell1.setCellValue(group);
                if (currentGroup == null) {
                    // Start new group
                    currentGroup = group;
                    groupStartCol = i;
                } else if (!currentGroup.equals(group)) {
                    // Group changed
                    // Merge previous group
                    if (groupStartCol < i - 1) {
                        sheet.addMergedRegion(new CellRangeAddress(0, 0, groupStartCol, i - 1));
                    }

                    // Start new group
                    currentGroup = group;
                    groupStartCol = i;
                }
                // If same group, continue
            }
        }

        // Close last group if exists
        if (currentGroup != null && groupStartCol < fields.size() - 1) {
            sheet.addMergedRegion(new CellRangeAddress(0, 0, groupStartCol, fields.size() - 1));
        }
    }

    private static void fillData(Sheet sheet, List<Field> fields, List<?> data, CellStyle dataStyle)
            throws IllegalAccessException {
        int rowNum = 2;
        for (Object item : data) {
            Row row = sheet.createRow(rowNum++);
            for (int i = 0; i < fields.size(); i++) {
                Field field = fields.get(i);
                field.setAccessible(true);
                Object value = field.get(item);
                ExcelField excelField = field.getAnnotation(ExcelField.class);

                Cell cell = row.createCell(i);
                cell.setCellStyle(dataStyle);

                if (value == null) {
                    cell.setCellValue("");
                } else if (value instanceof Date) {
                    // Simple date handling, better to use Local objects
                    cell.setCellValue(value.toString());
                } else if (value instanceof LocalDateTime) {
                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern(excelField.dateFormat());
                    cell.setCellValue(((LocalDateTime) value).format(formatter));
                } else if (value instanceof LocalDate) {
                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern(excelField.dateFormat());
                    cell.setCellValue(((LocalDate) value).format(formatter));
                } else if (value instanceof Number) {
                    cell.setCellValue(((Number) value).doubleValue());
                } else {
                    cell.setCellValue(value.toString());
                }
            }
        }
    }

    private static CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);

        setBorder(style);
        return style;
    }

    private static CellStyle createDataStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        setBorder(style);
        return style;
    }

    private static void setBorder(CellStyle style) {
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
    }
}
