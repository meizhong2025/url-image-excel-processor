package com.example.hairanalysis.client.dto;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;

import java.util.List;

@JsonIgnoreProperties(ignoreUnknown = true)
public class CardInfoResponse {
    private Integer code;
    private boolean success;
    private List<CardInfoDto> result;

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

    public List<CardInfoDto> getResult() {
        return result;
    }

    public void setResult(List<CardInfoDto> result) {
        this.result = result;
    }
}
