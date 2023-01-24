package org.smvs.bulkprocessor.service;

import org.smvs.bulkprocessor.model.AccountDetails;
import org.smvs.bulkprocessor.model.repository.AccountDetailsRepository;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.time.LocalDate;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

@Service
public class AccountReportService {
    private final AccountDetailsRepository accountDetailsRepository;
    private final ExcelUtils excelUtils;

    public AccountReportService(AccountDetailsRepository accountDetailsRepository, ExcelUtils excelUtils) {
        this.accountDetailsRepository = accountDetailsRepository;
        this.excelUtils = excelUtils;
    }

    public void addTempData() {
        List<AccountDetails> accountDetailsList = IntStream.range(1, 100)
                .mapToObj(index -> {
                    AccountDetails accountDetails = new AccountDetails();
                    accountDetails.setAccountNumber(100L + index);
                    accountDetails.setName("Testing: "+ index);
                    accountDetails.setActive(true);
                    accountDetails.setLocalDate(LocalDate.now());
                    accountDetails.setBalance(500L * index);
                    return accountDetails;
                }).collect(Collectors.toList());

        List<AccountDetails> savedDetails = accountDetailsRepository.saveAllAndFlush(accountDetailsList);
        System.out.println("Saved Data with ID:");
        savedDetails.forEach(accountDetails -> System.out.println(accountDetails.getId()));
    }

    public ByteArrayInputStream getAccountDetailsExcel(){
        List<AccountDetails> allAccountDetails = accountDetailsRepository.findAll();
        return excelUtils.accountDetailsToExcel(allAccountDetails);
    }
}
