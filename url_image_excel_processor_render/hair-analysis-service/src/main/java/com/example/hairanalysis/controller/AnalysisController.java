package com.example.hairanalysis.controller;

import com.example.hairanalysis.model.AnalysisRequest;
import com.example.hairanalysis.model.AnalysisResult;
import com.example.hairanalysis.service.AnalysisService;
import com.example.hairanalysis.service.ExcelExportService;
import jakarta.validation.Valid;
import org.springframework.http.ContentDisposition;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/api/analysis")
public class AnalysisController {

    private final AnalysisService analysisService;
    private final ExcelExportService excelExportService;

    public AnalysisController(AnalysisService analysisService, ExcelExportService excelExportService) {
        this.analysisService = analysisService;
        this.excelExportService = excelExportService;
    }

    @PostMapping
    public ResponseEntity<AnalysisResult> analyse(@Valid @RequestBody AnalysisRequest request) {
        AnalysisResult result = analysisService.analyseAndPersist(request);
        return ResponseEntity.ok(result);
    }

    @PostMapping("/export")
    public ResponseEntity<byte[]> export(@Valid @RequestBody AnalysisRequest request) {
        AnalysisResult result = analysisService.analyseAndPersist(request);
        byte[] excel = excelExportService.exportAnalysis(result);

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        headers.setContentDisposition(ContentDisposition.attachment().filename("analysis-" + result.getStoreId() + ".xlsx").build());
        headers.setContentLength(excel.length);
        return new ResponseEntity<>(excel, headers, HttpStatus.OK);
    }
}
