package com.example.hairanalysis.client;

import com.example.hairanalysis.client.dto.CardConsumptionRecordDto;
import com.example.hairanalysis.client.dto.CardConsumptionResponse;
import com.example.hairanalysis.client.dto.CardConsumptionResultDto;
import com.example.hairanalysis.client.dto.CardInfoDto;
import com.example.hairanalysis.client.dto.CardInfoResponse;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Component;
import org.springframework.util.CollectionUtils;
import org.springframework.web.client.RestTemplate;
import org.springframework.web.util.UriComponentsBuilder;

import java.math.BigDecimal;
import java.net.URI;
import java.security.MessageDigest;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.stream.Collectors;

@Component
public class BokaApiClient {

    private static final Logger log = LoggerFactory.getLogger(BokaApiClient.class);

    private final RestTemplate restTemplate;
    private final BokaApiProperties properties;

    public BokaApiClient(BokaApiProperties properties) {
        this.properties = properties;
        this.restTemplate = RestTemplateFactory.build(properties.getTimeoutMs());
    }

    public List<CardInfoDto> getCardsByPhone(String phone) {
        String timestamp = String.valueOf(Instant.now().toEpochMilli());
        HttpHeaders headers = buildHeaders(timestamp);
        URI uri = UriComponentsBuilder.fromHttpUrl(properties.getBaseUrl())
                .path("/s3connect/card/yidong/getCard")
                .queryParam("type", "Mphone")
                .queryParam("searchText", phone)
                .queryParam("status", "4,5")
                .build(true)
                .toUri();
        ResponseEntity<CardInfoResponse> response = restTemplate.exchange(uri, HttpMethod.GET, new HttpEntity<>(headers), CardInfoResponse.class);
        if (!response.hasBody()) {
            return Collections.emptyList();
        }
        CardInfoResponse body = response.getBody();
        if (body == null || !body.isSuccess() || body.getResult() == null) {
            log.warn("Failed to fetch card info for phone {} - response: {}", phone, body);
            return Collections.emptyList();
        }
        return body.getResult();
    }

    public List<CardConsumptionRecordDto> getConsumptionHistory(List<String> cardNos) {
        if (CollectionUtils.isEmpty(cardNos)) {
            return Collections.emptyList();
        }
        String cardNosParam = cardNos.stream().filter(Objects::nonNull).collect(Collectors.joining(","));
        if (cardNosParam.isEmpty()) {
            return Collections.emptyList();
        }
        int page = 1;
        int pageSize = 200;
        List<CardConsumptionRecordDto> result = new ArrayList<>();
        while (true) {
            String timestamp = String.valueOf(Instant.now().toEpochMilli());
            HttpHeaders headers = buildHeaders(timestamp);
            URI uri = UriComponentsBuilder.fromHttpUrl(properties.getBaseUrl())
                    .path("/s3connect/card/yidong/card/consume/history/list")
                    .queryParam("cardNos", cardNosParam)
                    .queryParam("page", page)
                    .queryParam("pageSize", pageSize)
                    .build(true)
                    .toUri();
            ResponseEntity<CardConsumptionResponse> response = restTemplate.exchange(uri, HttpMethod.GET, new HttpEntity<>(headers), CardConsumptionResponse.class);
            CardConsumptionResponse body = response.getBody();
            if (body == null || !body.isSuccess() || CollectionUtils.isEmpty(body.getResult())) {
                break;
            }
            Optional<CardConsumptionResultDto> maybeResult = body.getResult().stream().findFirst();
            if (maybeResult.isEmpty()) {
                break;
            }
            List<CardConsumptionRecordDto> pageList = maybeResult.get().getConsumptionList();
            if (!CollectionUtils.isEmpty(pageList)) {
                result.addAll(pageList);
            }
            if (pageList == null || pageList.size() < pageSize) {
                break;
            }
            page++;
        }
        return result;
    }

    private HttpHeaders buildHeaders(String timestamp) {
        HttpHeaders headers = new HttpHeaders();
        headers.add("appId", properties.getAppId());
        headers.add("timestamp", timestamp);
        headers.add("sign", sign(timestamp));
        headers.add(HttpHeaders.ACCEPT, "*/*");
        headers.add(HttpHeaders.USER_AGENT, "hair-analysis-service");
        return headers;
    }

    private String sign(String timestamp) {
        String raw = properties.getAppId() + properties.getAppKey() + timestamp;
        try {
            MessageDigest md5 = MessageDigest.getInstance("MD5");
            byte[] digest = md5.digest(raw.getBytes());
            StringBuilder builder = new StringBuilder();
            for (byte b : digest) {
                builder.append(String.format("%02x", b));
            }
            return builder.toString();
        } catch (Exception e) {
            log.error("Failed to sign request", e);
            throw new IllegalStateException("Unable to sign request", e);
        }
    }

    public BigDecimal sumConsumptionAmountForPrice(List<CardConsumptionRecordDto> consumptionRecords, BigDecimal targetPrice) {
        if (CollectionUtils.isEmpty(consumptionRecords) || targetPrice == null) {
            return BigDecimal.ZERO;
        }
        return consumptionRecords.stream()
                .map(CardConsumptionRecordDto::getConsumptionAmount)
                .filter(Objects::nonNull)
                .filter(amount -> amount.compareTo(targetPrice) == 0)
                .reduce(BigDecimal.ZERO, BigDecimal::add);
    }
}
