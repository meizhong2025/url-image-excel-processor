package com.example.hairanalysis.client.dto;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;

@JsonIgnoreProperties(ignoreUnknown = true)
public class CardInfoDto {

    @JsonProperty("cardCompId")
    private String cardCompId;
    @JsonProperty("cardMemberName")
    private String cardMemberName;
    @JsonProperty("cardName")
    private String cardName;
    @JsonProperty("cardNo")
    private String cardNo;
    @JsonProperty("cardPhone")
    private String cardPhone;
    @JsonProperty("cardType")
    private String cardType;
    @JsonProperty("status")
    private Integer status;

    public String getCardCompId() {
        return cardCompId;
    }

    public void setCardCompId(String cardCompId) {
        this.cardCompId = cardCompId;
    }

    public String getCardMemberName() {
        return cardMemberName;
    }

    public void setCardMemberName(String cardMemberName) {
        this.cardMemberName = cardMemberName;
    }

    public String getCardName() {
        return cardName;
    }

    public void setCardName(String cardName) {
        this.cardName = cardName;
    }

    public String getCardNo() {
        return cardNo;
    }

    public void setCardNo(String cardNo) {
        this.cardNo = cardNo;
    }

    public String getCardPhone() {
        return cardPhone;
    }

    public void setCardPhone(String cardPhone) {
        this.cardPhone = cardPhone;
    }

    public String getCardType() {
        return cardType;
    }

    public void setCardType(String cardType) {
        this.cardType = cardType;
    }

    public Integer getStatus() {
        return status;
    }

    public void setStatus(Integer status) {
        this.status = status;
    }
}
