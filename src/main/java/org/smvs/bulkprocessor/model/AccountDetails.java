package org.smvs.bulkprocessor.model;

import lombok.Data;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import java.time.LocalDate;

@Entity
@Data
public class AccountDetails {
    @Id
    @GeneratedValue(strategy = GenerationType.AUTO)
    private Integer id;

    @Column(name = "account_number")
    private Long accountNumber;

    @Column(name = "balance")
    private Long balance;

    @Column(name = "name")
    private String name;

    @Column(name = "start_date")
    private LocalDate localDate;

    @Column(name = "is_active")
    private boolean isActive;
}
