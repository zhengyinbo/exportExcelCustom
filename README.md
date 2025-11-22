# 自定义 Excel 导出 Demo

## 简介
本项目演示了如何在 Spring Boot 项目中使用自定义注解 `@ExportExcelCustom` 和 AOP 实现灵活的 Excel 导出功能。

## 功能特性
*   **自定义注解驱动**：只需在 Controller 方法上添加 `@ExportExcelCustom` 注解即可触发导出。
*   **灵活的字段配置**：通过 `@ExcelField` 注解配置列名、顺序、宽度、日期格式等。
*   **多级表头支持**：支持二级表头分组（一级表头）及自动合并。
*   **动态字段选择**：前端可以通过 `exportFields` 参数动态指定需要导出的字段。
*   **自动样式应用**：内置统一的表头和数据单元格样式（字体、边框、背景色）。

## 快速开始

### 1. 核心依赖
本项目基于 JDK 11 和 Spring Boot 2.7.x，主要依赖如下：
*   `spring-boot-starter-web`
*   `spring-boot-starter-aop`
*   `poi-ooxml` (Apache POI)

### 2. 使用方法

#### 步骤 1: 在 Controller 方法上添加注解
在返回 `List` 数据的接口上添加 `@ExportExcelCustom` 注解。
```java
@GetMapping("/users")
@ExportExcelCustom(fileName = "用户导出", sheetName = "用户数据")
public List<UserVO> exportUsers() {
    // 返回需要导出的数据列表
    return userService.findAll();
}
```

#### 步骤 2: 在数据模型 (VO) 上配置字段
使用 `@ExcelField` 注解定义 Excel 列的属性。
```java
public class UserVO {
    @ExcelField(title = "ID", order = 1, width = 10)
    private Long id;

    // 支持表头分组，group 为一级表头
    @ExcelField(title = "姓名", group = "基本信息", order = 2)
    private String name;

    @ExcelField(title = "邮箱", group = "联系方式", order = 4, width = 30)
    private String email;
    
    // 支持日期格式化
    @ExcelField(title = "注册日期", order = 6, dateFormat = "yyyy-MM-dd")
    private LocalDate registrationDate;
}
```

### 3. 前端调用示例

*   **导出所有配置的字段**：
    ```http
    GET http://localhost:8080/export/users
    ```

*   **导出指定字段** (使用 `exportFields` 参数，值为字段名，逗号分隔)：
    ```http
    GET http://localhost:8080/export/users?exportFields=name,email
    ```

## 项目结构
*   `annotation`: 自定义注解定义 (`ExportExcelCustom`, `ExcelField`)
*   `aspect`: AOP 切面处理 (`ExcelExportAspect`)
*   `utils`: Excel 生成工具类 (`ExcelGenerator`)
*   `controller`: 测试接口
*   `model`: 测试数据模型
