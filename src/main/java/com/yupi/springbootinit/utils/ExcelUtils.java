package com.yupi.springbootinit.utils;

import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.support.ExcelTypeEnum;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.ehcache.core.internal.statistics.StatsUtils;
import org.springframework.util.ResourceUtils;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * Excel 相关工具类
 */
@Slf4j
public class ExcelUtils {
    public static String excelToCsv(MultipartFile multipartFile){
        //File file = null;
        //try {
        //    file = ResourceUtils.getFile("/Users/air/code/project2/dubi-backend/dubi-backend/src/main/resources/网站数据.xlsx");
        //} catch (FileNotFoundException e) {
        //    throw new RuntimeException(e);
        //}
        List<Map<Integer, String>> list = null;
        //List<Map<Integer, String>> list = EasyExcel.read(file)
        //        .excelType(ExcelTypeEnum.XLSX)
        //        .sheet()
        //        .headRowNumber(0)
        //        .doReadSync();
        try {
            list = EasyExcel.read(multipartFile.getInputStream())
                    .excelType(ExcelTypeEnum.XLSX)
                    .sheet()
                    .headRowNumber(0)
                    .doReadSync();
        } catch (IOException e) {
            log.error("表格处理错误", e);
        }


        //如果数据为空，直接放返回""
        if(CollUtil.isEmpty(list)){
            return "";
        }
        //不为空 转换为csv
        StringBuilder stringBuilder = new StringBuilder();
        //读取表头
        LinkedHashMap<Integer, String> headerMap = (LinkedHashMap)list.get(0);
        List<String> headerList = headerMap.values().stream().filter(StringUtils::isNotEmpty).collect(Collectors.toList());
        System.out.println(StringUtils.join(headerList,"-"));
        stringBuilder.append(StringUtils.join(headerList,"-")).append("\n");
        //读取数据
        for (int i = 1; i < list.size(); i++) {
            LinkedHashMap<Integer, String> dataMap = (LinkedHashMap)list.get(i);
            List<String> dataList = dataMap.values().stream().filter(header -> StringUtils.isNotEmpty(header)).collect(Collectors.toList());
            System.out.println(StringUtils.join(dataList,"-"));
            stringBuilder.append(StringUtils.join(dataList,"-")).append("\n");

        }
        return stringBuilder.toString();

    }

    public static void main(String[] args) {
        excelToCsv(null);
    }
}
