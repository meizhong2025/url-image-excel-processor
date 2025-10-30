package com.example.hairanalysis.service;

import com.example.hairanalysis.client.BokaApiClient;
import com.example.hairanalysis.client.dto.CardConsumptionRecordDto;
import com.example.hairanalysis.mapper.boka.ConsumptionProjectRecordMapper;
import com.example.hairanalysis.mapper.boka.RechargeRecordMapper;
import com.example.hairanalysis.mapper.boka.SellCardDetailRecordMapper;
import com.example.hairanalysis.mapper.boka.SellCardRecordMapper;
import com.example.hairanalysis.mapper.mr.AnalysisFormMapper;
import com.example.hairanalysis.mapper.mr.ShopDataMapper;
import com.example.hairanalysis.model.AnalysisRequest;
import com.example.hairanalysis.model.AnalysisResult;
import com.example.hairanalysis.model.ConsumptionProjectRecord;
import com.example.hairanalysis.model.RechargeRecord;
import com.example.hairanalysis.model.SellCardDetailRecord;
import com.example.hairanalysis.model.SellCardRecord;
import com.example.hairanalysis.model.ShopData;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.util.CollectionUtils;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.YearMonth;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.stream.Collectors;

@Service
public class AnalysisService {

    private static final BigDecimal EXPERIENCE_PRICE = BigDecimal.valueOf(49);

    private final ShopDataMapper shopDataMapper;
    private final ConsumptionProjectRecordMapper consumptionProjectRecordMapper;
    private final SellCardRecordMapper sellCardRecordMapper;
    private final SellCardDetailRecordMapper sellCardDetailRecordMapper;
    private final RechargeRecordMapper rechargeRecordMapper;
    private final AnalysisFormMapper analysisFormMapper;
    private final BokaApiClient bokaApiClient;

    public AnalysisService(ShopDataMapper shopDataMapper,
                           ConsumptionProjectRecordMapper consumptionProjectRecordMapper,
                           SellCardRecordMapper sellCardRecordMapper,
                           SellCardDetailRecordMapper sellCardDetailRecordMapper,
                           RechargeRecordMapper rechargeRecordMapper,
                           AnalysisFormMapper analysisFormMapper,
                           BokaApiClient bokaApiClient) {
        this.shopDataMapper = shopDataMapper;
        this.consumptionProjectRecordMapper = consumptionProjectRecordMapper;
        this.sellCardRecordMapper = sellCardRecordMapper;
        this.sellCardDetailRecordMapper = sellCardDetailRecordMapper;
        this.rechargeRecordMapper = rechargeRecordMapper;
        this.analysisFormMapper = analysisFormMapper;
        this.bokaApiClient = bokaApiClient;
    }

    @Transactional
    public AnalysisResult analyseAndPersist(AnalysisRequest request) {
        YearMonth yearMonth = YearMonth.parse(request.getMonth());
        LocalDate monthStart = yearMonth.atDay(1);
        LocalDate monthEnd = yearMonth.atEndOfMonth();
        LocalDateTime monthStartDateTime = monthStart.atStartOfDay();
        LocalDateTime monthEndDateTime = monthEnd.atTime(23, 59, 59);

        ShopData shop = Optional.ofNullable(shopDataMapper.findById(request.getStoreId()))
                .orElseThrow(() -> new IllegalArgumentException("Store not found or disabled: " + request.getStoreId()));

        AnalysisResult result = new AnalysisResult();
        result.setStoreId(shop.getId());
        result.setStoreName(shop.getName());
        result.setDepartmentName(shop.getPname());
        result.setMonth(request.getMonth());

        List<ConsumptionProjectRecord> consumptionRecords = consumptionProjectRecordMapper.findByStoreAndMonth(shop.getNum(), monthStartDateTime, monthEndDateTime);
        computeExperienceMetrics(consumptionRecords, result);

        List<SellCardRecord> sellCardRecords = sellCardRecordMapper.findByStoreAndMonth(shop.getNum(), monthStart, monthEnd);
        Map<String, SellCardRecord> sellCardRecordByBill = sellCardRecords.stream().collect(Collectors.toMap(SellCardRecord::getBillId, r -> r, (r1, r2) -> r1));

        List<SellCardDetailRecord> sellCardDetailRecords = sellCardDetailRecordMapper.findByStoreAndMonth(shop.getNum(), monthStart, monthEnd);
        List<RechargeRecord> rechargeRecords = rechargeRecordMapper.findByStoreAndMonth(shop.getNum(), monthStart, monthEnd);

        computeDealMetrics(result, sellCardDetailRecords, sellCardRecordByBill, rechargeRecords, monthStart);

        analysisFormMapper.insert(result);
        return result;
    }

    private void computeExperienceMetrics(List<ConsumptionProjectRecord> consumptionRecords, AnalysisResult result) {
        if (CollectionUtils.isEmpty(consumptionRecords)) {
            return;
        }
        Set<String> experienceOrderIds = new HashSet<>();
        BigDecimal experienceAmount = BigDecimal.ZERO;
        for (ConsumptionProjectRecord record : consumptionRecords) {
            if (record.getPrice() == null || EXPERIENCE_PRICE.compareTo(record.getPrice()) != 0) {
                continue;
            }
            if (record.getBillId() != null) {
                experienceOrderIds.add(record.getBillId());
            }
            if (record.getTotalPrice() != null) {
                experienceAmount = experienceAmount.add(record.getTotalPrice());
            }
        }
        result.setHairExperienceCount(experienceOrderIds.size());
        result.setHairExperienceAmount(experienceAmount);
    }

    private void computeDealMetrics(AnalysisResult result,
                                    List<SellCardDetailRecord> sellCardDetailRecords,
                                    Map<String, SellCardRecord> sellCardRecordByBill,
                                    List<RechargeRecord> rechargeRecords,
                                    LocalDate monthStart) {
        Map<String, List<CardConsumptionRecordDto>> consumptionCache = new HashMap<>();
        int dealCount = 0;
        BigDecimal dealAmount = BigDecimal.ZERO;

        if (!CollectionUtils.isEmpty(sellCardDetailRecords)) {
            for (SellCardDetailRecord detail : sellCardDetailRecords) {
                SellCardRecord master = sellCardRecordByBill.get(detail.getBillId());
                if (master == null || master.getMobile() == null || detail.getMemberCardNo() == null) {
                    continue;
                }
                if (!isFirstCardForMobile(master.getMobile(), master)) {
                    continue;
                }
                if (!hasExperienceRecord(detail.getMemberCardNo(), master.getCompId(), consumptionCache)) {
                    continue;
                }
                dealCount++;
                BigDecimal amount = Optional.ofNullable(detail.getStoredAmount())
                        .map(BigDecimal::valueOf)
                        .orElseGet(() -> Optional.ofNullable(master.getPaidAmount()).map(BigDecimal::valueOf).orElse(BigDecimal.ZERO));
                dealAmount = dealAmount.add(amount);
            }
        }

        if (!CollectionUtils.isEmpty(rechargeRecords)) {
            for (RechargeRecord recharge : rechargeRecords) {
                if (recharge.getMemberCardNo() == null) {
                    continue;
                }
                if (!isFirstRecharge(recharge.getMemberCardNo(), monthStart)) {
                    continue;
                }
                if (!hasExperienceRecord(recharge.getMemberCardNo(), recharge.getCompId(), consumptionCache)) {
                    continue;
                }
                dealCount++;
                dealAmount = dealAmount.add(Optional.ofNullable(recharge.getStoredAmount()).map(BigDecimal::valueOf).orElse(BigDecimal.ZERO));
            }
        }

        result.setHairDealCount(dealCount);
        result.setHairDealAmount(dealAmount);
    }

    private boolean isFirstCardForMobile(String mobile, SellCardRecord currentRecord) {
        List<SellCardRecord> earliest = sellCardRecordMapper.findFirstSellRecordByMobile(mobile);
        if (CollectionUtils.isEmpty(earliest)) {
            return true;
        }
        SellCardRecord firstRecord = earliest.get(0);
        return Objects.equals(firstRecord.getBillId(), currentRecord.getBillId());
    }

    private boolean isFirstRecharge(String cardNo, LocalDate monthStart) {
        List<RechargeRecord> history = rechargeRecordMapper.findRechargeHistory(cardNo, monthStart);
        return CollectionUtils.isEmpty(history);
    }

    private boolean hasExperienceRecord(String cardNo, String compId, Map<String, List<CardConsumptionRecordDto>> consumptionCache) {
        List<CardConsumptionRecordDto> records = consumptionCache.computeIfAbsent(cardNo, key -> bokaApiClient.getConsumptionHistory(List.of(key)));
        if (CollectionUtils.isEmpty(records)) {
            return false;
        }
        return records.stream()
                .filter(record -> record.getConsumptionAmount() != null && EXPERIENCE_PRICE.compareTo(record.getConsumptionAmount()) == 0)
                .anyMatch(record -> compId == null || compId.equals(record.getShopId()));
    }
}
