package com.example.hairanalysis.model;

import jakarta.validation.constraints.NotBlank;
import jakarta.validation.constraints.Pattern;

public class AnalysisRequest {

    @NotBlank
    private String storeId;

    @NotBlank
    @Pattern(regexp = "\\d{4}-\\d{2}", message = "month must be in the format yyyy-MM")
    private String month;

    public String getStoreId() {
        return storeId;
    }

    public void setStoreId(String storeId) {
        this.storeId = storeId;
    }

    public String getMonth() {
        return month;
    }

    public void setMonth(String month) {
        this.month = month;
    }
}
