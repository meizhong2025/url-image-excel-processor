package com.example.hairanalysis.client.dto;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;

import java.util.List;

@JsonIgnoreProperties(ignoreUnknown = true)
public class CardConsumptionResponse {
    private Integer code;
    private boolean success;
    private List<CardConsumptionResultDto> result;

    public Integer getCode() {
        return code;
    }

    public void setCode(Integer code) {
        this.code = code;
    }

    public boolean isSuccess() {
        return success;
    }

    public void setSuccess(boolean success) {
        this.success = success;
    }

    public List<CardConsumptionResultDto> getResult() {
        return result;
    }

    public void setResult(List<CardConsumptionResultDto> result) {
        this.result = result;
    }
}
