package org.smvs.bulkprocessor.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.smvs.bulkprocessor.model.AccountDetails;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

@Service
public class ExcelUtils {
    private static final String[] HEADERs = {"Id", "Account Number", "Account Name", "Start Date", "Balance"};

    public ByteArrayInputStream accountDetailsToExcel(List<AccountDetails> accountDetails) {

        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream();) {
            Sheet sheet = workbook.createSheet("AccountDetails");

            // Header
            Row headerRow = sheet.createRow(0);

            for (int col = 0; col < HEADERs.length; col++) {
                Cell cell = headerRow.createCell(col);
                cell.setCellValue(HEADERs[col]);
            }

            int rowIdx = 1;
            for (AccountDetails details : accountDetails) {
                Row row = sheet.createRow(rowIdx++);

                row.createCell(0).setCellValue(details.getId());
                row.createCell(1).setCellValue(details.getAccountNumber());
                row.createCell(2).setCellValue(details.getName());
                row.createCell(3).setCellValue(details.getLocalDate());
                row.createCell(4).setCellValue(details.getBalance());
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            throw new RuntimeException("fail to import data to Excel file: " + e.getMessage());
        }
    }
}
