package com.excel.download.service.abstarct;


import lombok.Getter;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.ibatis.session.ResultContext;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.HashMap;
import java.util.Map;
import java.util.Objects;


/*
 *
 *  엑셀 다운로드 공통 추상화 객체
 *
 */
@Slf4j
@Getter
public abstract class AbstractExcelDown {

    protected static final int MAX_EXCEL_ROW_SIZE = 60000;
    private Map<String, CellStyle> styleMap; // 스타일맵
    private int rows; // cell row 위치
    private SXSSFWorkbook workbook; // 엑셀
    private SXSSFSheet sheet; // 엑셀-시트

    protected abstract String getName(); // Sheet 이름 가져오기


    protected SXSSFRow createRow(int a) {
        SXSSFRow sXSSFRow = null;
        for (int i = 0; i < a; i++) {
            rows++;
            sXSSFRow = sheet.createRow(rows);
        }
        return sXSSFRow;
    }

    protected abstract void write(ResultContext<?> resultContext); // Data write

    private void createWorkbook() {
        // SXSSFWorkbook 생성
        SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(1000); // 메모리 행 10000개로 제한, 초과 시 Disk로 flush
        sxssfWorkbook.setCompressTempFiles(true); //임시파일 압축
        workbook = sxssfWorkbook;
        styleMap = setContentsStyle();
    }

    /**
     * contnet의 style 설정
     *
     * @return Map<String, CellStyle>
     */
    private Map<String, CellStyle> setContentsStyle() {

        Map<String, CellStyle> contentMap = new HashMap<>();

        String[] alignList = new String[]{"left", "center", "right"};

        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 10);
        font.setFontName("맑은고딕");
        font.setBold(false);
        for (String align : alignList) {
            CellStyle style = workbook.createCellStyle();
            style.setFont(font);
            style.setVerticalAlignment(VerticalAlignment.CENTER);
            switch (align) {
                case "left":
                    style.setAlignment(HorizontalAlignment.LEFT);
                    break;
                case "right":
                    style.setAlignment(HorizontalAlignment.RIGHT);
                    break;
                default:
                    style.setAlignment(HorizontalAlignment.CENTER);
            }
            contentMap.put(align, style);
        }
        return contentMap;
    }

    public void createSheet() {
        rows = 0;
        String sheetName = getName();
        sheetName = workbook.getNumberOfSheets() > 0 ? sheetName + "_" + StringUtils.defaultString(String.valueOf(workbook.getNumberOfSheets()), "") : sheetName;
        SXSSFSheet sxssfSheet = workbook.createSheet(sheetName);
        sxssfSheet.setRandomAccessWindowSize(1000); // 메모리 행 10000개로 제한, 초과 시 Disk로 flush
        sxssfSheet.setDefaultColumnWidth(9);         // Cell 스타일 값
        //column width 설정
        sheet = sxssfSheet;
    } // Sheet 생성

    // 비즈니스 실행
    public void work() {
        log.info("ExcelDownload Target Service Name : {}", getName());
        createWorkbook(); // 워크북 생성
        createSheet(); // 시트생성
        setBody(); // Data write
    }
    protected abstract void setBody(); // 데이터 생성
    protected abstract String getFileName(); //파일 이름

    protected void createCell(SXSSFRow row, int cellNo, Object value) {
        createCell(row, cellNo, value, CellType.STRING, "center");
    }

    protected void createCell(SXSSFRow row, int cellNo, Object value, CellType cellType, String cellStyle) {
        SXSSFCell cell = row.createCell(cellNo, cellType);
        cell.setCellValue(Objects.toString(value, "-"));
        cell.setCellStyle(getStyleMap().get(cellStyle));
    }


}
