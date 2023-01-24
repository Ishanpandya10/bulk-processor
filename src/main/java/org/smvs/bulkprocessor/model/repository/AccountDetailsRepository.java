package org.smvs.bulkprocessor.model.repository;

import org.smvs.bulkprocessor.model.AccountDetails;
import org.springframework.data.jpa.repository.JpaRepository;

public interface AccountDetailsRepository extends JpaRepository<AccountDetails, Integer> {
}
