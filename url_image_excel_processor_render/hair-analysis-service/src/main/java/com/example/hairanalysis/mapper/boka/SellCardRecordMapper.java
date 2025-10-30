package com.example.hairanalysis.mapper.boka;

import com.example.hairanalysis.model.SellCardRecord;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.time.LocalDate;
import java.util.List;

@Mapper
public interface SellCardRecordMapper {

    List<SellCardRecord> findByStoreAndMonth(@Param("storeId") String storeId,
                                             @Param("startDate") LocalDate startDate,
                                             @Param("endDate") LocalDate endDate);

    List<SellCardRecord> findFirstSellRecordByMobile(@Param("mobile") String mobile);
}
