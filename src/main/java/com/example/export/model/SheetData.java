package com.example.export.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;
import java.util.Set;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class SheetData {
    /**
     * The name of the sheet.
     */
    private String name;

    /**
     * The data list for this sheet.
     */
    private List<?> data;

    /**
     * The class type of the data elements.
     */
    private Class<?> clazz;

    /**
     * Optional: Specific fields to export for this sheet.
     */
    private Set<String> exportFields;

    public SheetData(String name, List<?> data, Class<?> clazz) {
        this.name = name;
        this.data = data;
        this.clazz = clazz;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<?> getData() {
        return data;
    }

    public void setData(List<?> data) {
        this.data = data;
    }

    public Class<?> getClazz() {
        return clazz;
    }

    public void setClazz(Class<?> clazz) {
        this.clazz = clazz;
    }

    public Set<String> getExportFields() {
        return exportFields;
    }

    public void setExportFields(Set<String> exportFields) {
        this.exportFields = exportFields;
    }
}
