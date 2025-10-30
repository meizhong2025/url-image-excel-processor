package com.example.hairanalysis.mapper.boka;

import com.example.hairanalysis.model.SellCardDetailRecord;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.time.LocalDate;
import java.util.List;

@Mapper
public interface SellCardDetailRecordMapper {

    List<SellCardDetailRecord> findByStoreAndMonth(@Param("storeId") String storeId,
                                                   @Param("startDate") LocalDate startDate,
                                                   @Param("endDate") LocalDate endDate);

    List<SellCardDetailRecord> findByBillId(@Param("billId") String billId);
}
