package com.example.export.controller;

import com.example.export.annotation.ExportExcelCustom;
import com.example.export.model.UserVO;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

@RestController
@RequestMapping("/export")
public class TestController {

    @GetMapping("/users")
    @ExportExcelCustom(fileName = "user_export", sheetName = "User Data")
    public List<UserVO> exportUsers() {
        List<UserVO> users = new ArrayList<>();
        users.add(new UserVO(1L, "Alice", 30, "alice@example.com", "1234567890", LocalDate.now().minusDays(10),
                LocalDateTime.now().minusHours(2)));
        users.add(new UserVO(2L, "Bob", 25, "bob@example.com", "0987654321", LocalDate.now().minusDays(5),
                LocalDateTime.now().minusHours(5)));
        users.add(new UserVO(3L, "Charlie", 35, "charlie@example.com", "1122334455", LocalDate.now().minusDays(20),
                LocalDateTime.now().minusHours(10)));
        return users;
    }
}
