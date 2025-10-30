package com.example.hairanalysis.model;

import java.math.BigDecimal;

public class AnalysisResult {

    private String departmentName;
    private String storeId;
    private String storeName;
    private String month;

    private int hairExperienceCount;
    private BigDecimal hairExperienceAmount = BigDecimal.ZERO;
    private int hairDealCount;
    private BigDecimal hairDealAmount = BigDecimal.ZERO;

    private int groundPushExperienceCount;
    private BigDecimal groundPushExperienceAmount = BigDecimal.ZERO;
    private int groundPushDealCount;
    private BigDecimal groundPushDealAmount = BigDecimal.ZERO;

    public int getTotalExperienceCount() {
        return hairExperienceCount + groundPushExperienceCount;
    }

    public BigDecimal getTotalExperienceAmount() {
        return hairExperienceAmount.add(groundPushExperienceAmount);
    }

    public int getTotalDealCount() {
        return hairDealCount + groundPushDealCount;
    }

    public BigDecimal getTotalDealAmount() {
        return hairDealAmount.add(groundPushDealAmount);
    }

    public String getDepartmentName() {
        return departmentName;
    }

    public void setDepartmentName(String departmentName) {
        this.departmentName = departmentName;
    }

    public String getStoreId() {
        return storeId;
    }

    public void setStoreId(String storeId) {
        this.storeId = storeId;
    }

    public String getStoreName() {
        return storeName;
    }

    public void setStoreName(String storeName) {
        this.storeName = storeName;
    }

    public String getMonth() {
        return month;
    }

    public void setMonth(String month) {
        this.month = month;
    }

    public int getHairExperienceCount() {
        return hairExperienceCount;
    }

    public void setHairExperienceCount(int hairExperienceCount) {
        this.hairExperienceCount = hairExperienceCount;
    }

    public BigDecimal getHairExperienceAmount() {
        return hairExperienceAmount;
    }

    public void setHairExperienceAmount(BigDecimal hairExperienceAmount) {
        this.hairExperienceAmount = hairExperienceAmount;
    }

    public int getHairDealCount() {
        return hairDealCount;
    }

    public void setHairDealCount(int hairDealCount) {
        this.hairDealCount = hairDealCount;
    }

    public BigDecimal getHairDealAmount() {
        return hairDealAmount;
    }

    public void setHairDealAmount(BigDecimal hairDealAmount) {
        this.hairDealAmount = hairDealAmount;
    }

    public int getGroundPushExperienceCount() {
        return groundPushExperienceCount;
    }

    public void setGroundPushExperienceCount(int groundPushExperienceCount) {
        this.groundPushExperienceCount = groundPushExperienceCount;
    }

    public BigDecimal getGroundPushExperienceAmount() {
        return groundPushExperienceAmount;
    }

    public void setGroundPushExperienceAmount(BigDecimal groundPushExperienceAmount) {
        this.groundPushExperienceAmount = groundPushExperienceAmount;
    }

    public int getGroundPushDealCount() {
        return groundPushDealCount;
    }

    public void setGroundPushDealCount(int groundPushDealCount) {
        this.groundPushDealCount = groundPushDealCount;
    }

    public BigDecimal getGroundPushDealAmount() {
        return groundPushDealAmount;
    }

    public void setGroundPushDealAmount(BigDecimal groundPushDealAmount) {
        this.groundPushDealAmount = groundPushDealAmount;
    }
}
