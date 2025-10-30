package com.example.hairanalysis.service;

import com.example.hairanalysis.model.AnalysisResult;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

@Service
public class ExcelExportService {

    public byte[] exportAnalysis(AnalysisResult result) {
        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("分析结果");
            sheet.setDefaultColumnWidth(18);

            CellStyle titleStyle = workbook.createCellStyle();
            titleStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
            titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            titleStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
            XSSFFont titleFont = ((XSSFWorkbook) workbook).createFont();
            titleFont.setBold(true);
            titleFont.setFontHeightInPoints((short) 14);
            titleStyle.setFont(titleFont);

            Row titleRow = sheet.createRow(0);
            titleRow.setHeightInPoints(24);
            Cell titleCell = titleRow.createCell(0);
            titleCell.setCellValue("美容月份新客来源及成交分析");
            titleCell.setCellStyle(titleStyle);
            sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(0, 0, 0, 10));

            Row headerRow = sheet.createRow(1);
            headerRow.createCell(0).setCellValue("所属部");
            headerRow.createCell(1).setCellValue("门店");
            headerRow.createCell(2).setCellValue("月份");
            headerRow.createCell(3).setCellValue("美发体验人数");
            headerRow.createCell(4).setCellValue("美发体验业绩");
            headerRow.createCell(5).setCellValue("美发成交人数");
            headerRow.createCell(6).setCellValue("美发成交业绩");
            headerRow.createCell(7).setCellValue("地拓体验人数");
            headerRow.createCell(8).setCellValue("地拓体验业绩");
            headerRow.createCell(9).setCellValue("地拓成交人数");
            headerRow.createCell(10).setCellValue("地拓成交业绩");

            Row dataRow = sheet.createRow(2);
            dataRow.createCell(0).setCellValue(result.getDepartmentName());
            dataRow.createCell(1).setCellValue(result.getStoreName());
            dataRow.createCell(2).setCellValue(result.getMonth());
            dataRow.createCell(3).setCellValue(result.getHairExperienceCount());
            dataRow.createCell(4).setCellValue(result.getHairExperienceAmount().doubleValue());
            dataRow.createCell(5).setCellValue(result.getHairDealCount());
            dataRow.createCell(6).setCellValue(result.getHairDealAmount().doubleValue());
            dataRow.createCell(7).setCellValue(result.getGroundPushExperienceCount());
            dataRow.createCell(8).setCellValue(result.getGroundPushExperienceAmount().doubleValue());
            dataRow.createCell(9).setCellValue(result.getGroundPushDealCount());
            dataRow.createCell(10).setCellValue(result.getGroundPushDealAmount().doubleValue());

            workbook.write(outputStream);
            return outputStream.toByteArray();
        } catch (IOException e) {
            throw new IllegalStateException("Failed to export analysis excel", e);
        }
    }
}
