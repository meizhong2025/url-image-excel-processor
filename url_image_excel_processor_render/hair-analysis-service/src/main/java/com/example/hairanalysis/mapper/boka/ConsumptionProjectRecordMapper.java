package com.example.hairanalysis.mapper.boka;

import com.example.hairanalysis.model.ConsumptionProjectRecord;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.time.LocalDateTime;
import java.util.List;

@Mapper
public interface ConsumptionProjectRecordMapper {

    List<ConsumptionProjectRecord> findByStoreAndMonth(@Param("storeId") String storeId,
                                                       @Param("startDate") LocalDateTime startDate,
                                                       @Param("endDate") LocalDateTime endDate);

    List<ConsumptionProjectRecord> findByBillId(@Param("billId") String billId);
}
