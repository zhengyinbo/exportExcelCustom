package com.example.export.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Custom annotation to mark methods for Excel export.
 */
@Target(ElementType.METHOD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExportExcelCustom {

    /**
     * The name of the file to be exported (without extension).
     */
    String fileName() default "export";

    /**
     * The name of the sheet in the Excel file.
     */
    String sheetName() default "Sheet1";
}
