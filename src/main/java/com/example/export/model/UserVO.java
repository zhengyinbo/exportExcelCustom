package com.example.export.model;

import com.example.export.annotation.ExcelField;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDate;
import java.time.LocalDateTime;

@Data
@NoArgsConstructor
public class UserVO {

    @ExcelField(title = "ID", order = 1, width = 10)
    private Long id;

    @ExcelField(title = "Name", group = "Basic Info", order = 2)
    private String name;

    @ExcelField(title = "Age", group = "Basic Info", order = 3, width = 10)
    private Integer age;

    @ExcelField(title = "Email", group = "Contact", order = 4, width = 30)
    private String email;

    @ExcelField(title = "Phone", group = "Contact", order = 5, width = 20)
    private String phone;

    @ExcelField(title = "Registration Date", order = 6, width = 25, dateFormat = "yyyy-MM-dd")
    private LocalDate registrationDate;

    @ExcelField(title = "Last Login", order = 7, width = 25)
    private LocalDateTime lastLogin;

    public UserVO(Long id, String name, Integer age, String email, String phone, LocalDate registrationDate,
            LocalDateTime lastLogin) {
        this.id = id;
        this.name = name;
        this.age = age;
        this.email = email;
        this.phone = phone;
        this.registrationDate = registrationDate;
        this.lastLogin = lastLogin;
    }
}
