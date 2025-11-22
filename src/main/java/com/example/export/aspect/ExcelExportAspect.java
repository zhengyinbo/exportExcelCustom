package com.example.export.aspect;

import com.example.export.annotation.ExportExcelCustom;
import com.example.export.utils.ExcelGenerator;
import org.apache.poi.ss.usermodel.Workbook;
import org.aspectj.lang.ProceedingJoinPoint;
import org.aspectj.lang.annotation.Around;
import org.aspectj.lang.annotation.Aspect;
import org.springframework.stereotype.Component;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.Arrays;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

@Aspect
@Component
public class ExcelExportAspect {

    @Around("@annotation(exportExcelCustom)")
    public Object around(ProceedingJoinPoint point, ExportExcelCustom exportExcelCustom) throws Throwable {
        Object result = point.proceed();

        if (result instanceof List) {
            List<?> data = (List<?>) result;
            if (!data.isEmpty()) {
                ServletRequestAttributes attributes = (ServletRequestAttributes) RequestContextHolder
                        .getRequestAttributes();
                if (attributes != null) {
                    HttpServletRequest request = attributes.getRequest();
                    String exportFieldsParam = request.getParameter("exportFields");
                    Set<String> exportFields = null;
                    if (exportFieldsParam != null && !exportFieldsParam.isEmpty()) {
                        exportFields = new HashSet<>(Arrays.asList(exportFieldsParam.split(",")));
                    }

                    Class<?> clazz = data.get(0).getClass();
                    Workbook workbook = ExcelGenerator.generate(data, clazz, exportExcelCustom.sheetName(),
                            exportFields);

                    HttpServletResponse response = attributes.getResponse();
                    if (response != null) {
                        String fileName = URLEncoder.encode(exportExcelCustom.fileName() + ".xlsx",
                                StandardCharsets.UTF_8.toString());
                        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                        response.setHeader("Content-Disposition", "attachment; filename=\"" + fileName + "\"");

                        workbook.write(response.getOutputStream());
                        workbook.close();
                        return null; // Response handled
                    }
                }
            }
        }
        return result;
    }
}
