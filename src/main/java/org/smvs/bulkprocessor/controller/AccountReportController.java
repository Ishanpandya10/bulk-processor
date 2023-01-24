package org.smvs.bulkprocessor.controller;

import org.smvs.bulkprocessor.service.AccountReportService;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping(path = "/account-report")
public class AccountReportController {
    private final AccountReportService accountReportService;

    public AccountReportController(AccountReportService accountReportService) {
        this.accountReportService = accountReportService;
    }

    @GetMapping("/hello")
    public String hello() {
        return "Hi There";
    }

    @GetMapping("/add-temp-data")
    public String addTempData() {
        accountReportService.addTempData();
        return "Done adding data";
    }

    @GetMapping("/get-account-details-excel")
    public ResponseEntity<Resource> getAccountDetailsExcel() {
        String filename = "account_details.xlsx";
        InputStreamResource file = new InputStreamResource(accountReportService.getAccountDetailsExcel());

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                .contentType(MediaType.parseMediaType("application/vnd.ms-excel"))
                .body(file);
    }
}