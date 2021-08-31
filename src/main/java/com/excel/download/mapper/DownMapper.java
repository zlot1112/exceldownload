package com.excel.download.mapper;

import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.session.ResultHandler;

import java.util.Map;

@Mapper
public interface DownMapper {
    void get(ResultHandler<Map<String, Object>> resultHandler);
}
