package com.excel.download.view;

import com.excel.download.mapper.DownMapper;
import com.excel.download.service.DownService;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;
import org.springframework.web.servlet.view.AbstractView;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.Map;

@Slf4j
@Component("ExcelDownView")
public class ExcelDownView extends AbstractView {

    private final DownMapper downMapper;

    ExcelDownView(DownMapper downMapper) {
        this.downMapper = downMapper;
    }

    @Override
    protected void renderMergedOutputModel(Map<String, Object> model, HttpServletRequest request, HttpServletResponse response) throws Exception {
        ServletOutputStream out = null;

        DownService downService =
                new DownService(downMapper);
        try {

            downService.work();
            //Response write
            out = response.getOutputStream();
            setContentType("application/vnd.ms-excel"); // set excel content type
            response.setHeader("Content-Disposition", "attachment; filename=\"" + downService.getFileName() + "\"");
            downService.getWorkbook().write(out);

        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        } finally {
            try {
                if (downService.getWorkbook() != null) {
                    downService.getWorkbook().close();
                    downService.getWorkbook().dispose();
                }
                if (out != null) out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
