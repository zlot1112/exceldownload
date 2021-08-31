package com.excel.download.service;


import com.excel.download.mapper.DownMapper;
import com.excel.download.service.abstarct.AbstractExcelDown;
import lombok.extern.slf4j.Slf4j;
import org.apache.ibatis.session.ResultContext;
import org.apache.ibatis.session.ResultHandler;
import org.apache.poi.xssf.streaming.SXSSFRow;

import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;


/**
 * Excel Download 구현 클래스
 *
 * @version 1.0
 */
@Slf4j
public class DownService extends AbstractExcelDown {

    private final DownMapper downMapper;
    private int rowNum;

    public DownService(DownMapper downMapper) {
        this.downMapper = downMapper;
    }
    @Override
    public String getName() {
        return "ExcelDownLoad Test";
    }


    @Override
    public void setBody() {
        // 데이터 작성영역
        ResultHandler<Map<String,Object>> resultHandler = this::write;
        downMapper.get(resultHandler);
    }

    @Override
    protected void write(ResultContext<?> resultContext) {

        Map<String,Object> map = (Map<String,Object>) resultContext.getResultObject();
        rowNum++;
        SXSSFRow row = createRow(1);
        createCell(row, 1, String.valueOf(rowNum)); // No.
        AtomicInteger cellNo = new AtomicInteger(1);

        map.forEach((k,v)-> createCell(row, cellNo.getAndIncrement(), k  +"_"+ v));
        //6만건 이상 ROW 발생시 끊기
        if (getRows() > MAX_EXCEL_ROW_SIZE) {
            createSheet(); // 시트 재 생성
        }
    }

    @Override
    public String getFileName() {
        return getName();
    }

}
