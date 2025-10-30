package com.example.hairanalysis.client.dto;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;

import java.util.List;

@JsonIgnoreProperties(ignoreUnknown = true)
public class CardConsumptionResultDto {

    @JsonProperty("cardNo")
    private String cardNo;
    @JsonProperty("consumptionList")
    private List<CardConsumptionRecordDto> consumptionList;

    public String getCardNo() {
        return cardNo;
    }

    public void setCardNo(String cardNo) {
        this.cardNo = cardNo;
    }

    public List<CardConsumptionRecordDto> getConsumptionList() {
        return consumptionList;
    }

    public void setConsumptionList(List<CardConsumptionRecordDto> consumptionList) {
        this.consumptionList = consumptionList;
    }
}
