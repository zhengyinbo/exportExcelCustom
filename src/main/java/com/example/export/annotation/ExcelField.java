package com.example.export.annotation;

import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Annotation to configure Excel fields.
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelField {

    /**
     * Level 2 header name (the actual column name).
     */
    String title();

    /**
     * Level 1 header name (for grouping). If empty, no grouping.
     */
    String group() default "";

    /**
     * Order of the column.
     */
    int order() default 0;

    /**
     * Column width.
     */
    int width() default 20;

    /**
     * Background color for the header (IndexedColors name).
     * Default is GREY_25_PERCENT.
     */
    IndexedColors headerColor() default IndexedColors.GREY_25_PERCENT;

    /**
     * Date format for date fields.
     */
    String dateFormat() default "yyyy-MM-dd HH:mm:ss";
}
