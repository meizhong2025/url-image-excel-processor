package com.example.hairanalysis.mapper.boka;

import com.example.hairanalysis.model.RechargeRecord;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.time.LocalDate;
import java.util.List;

@Mapper
public interface RechargeRecordMapper {

    List<RechargeRecord> findByStoreAndMonth(@Param("storeId") String storeId,
                                             @Param("startDate") LocalDate startDate,
                                             @Param("endDate") LocalDate endDate);

    List<RechargeRecord> findRechargeHistory(@Param("cardNo") String cardNo,
                                             @Param("beforeDate") LocalDate beforeDate);
}
