package com.example.hairanalysis.model;

import java.time.LocalDate;

public class RechargeRecord {
    private Long id;
    private String compId;
    private String billId;
    private LocalDate billDate;
    private String memberCardNo;
    private Double storedAmount;
    private String mobile;

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getCompId() {
        return compId;
    }

    public void setCompId(String compId) {
        this.compId = compId;
    }

    public String getBillId() {
        return billId;
    }

    public void setBillId(String billId) {
        this.billId = billId;
    }

    public LocalDate getBillDate() {
        return billDate;
    }

    public void setBillDate(LocalDate billDate) {
        this.billDate = billDate;
    }

    public String getMemberCardNo() {
        return memberCardNo;
    }

    public void setMemberCardNo(String memberCardNo) {
        this.memberCardNo = memberCardNo;
    }

    public Double getStoredAmount() {
        return storedAmount;
    }

    public void setStoredAmount(Double storedAmount) {
        this.storedAmount = storedAmount;
    }

    public String getMobile() {
        return mobile;
    }

    public void setMobile(String mobile) {
        this.mobile = mobile;
    }
}
