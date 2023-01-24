package com.example.webtest.controller;

import io.micrometer.common.util.StringUtils;
import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.log4j.Log4j2;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.IndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import java.awt.*;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

@Controller
@RequestMapping("/")
@Log4j2
public class testController {

    @GetMapping("/test1")
    public void test1(HttpServletResponse response) throws IOException {
        this.b2bOutGoodsCheck(response);
        String test = "";
    }

    public static void b2bOutGoodsCheck(HttpServletResponse response) throws IOException {
        // 워크북 생성
        SXSSFWorkbook workbook = new SXSSFWorkbook();


        Cell cell = null;
        Row row0 = null;
        Row row1 = null;
        Row row2 = null;
        Row row3 = null;
        Row row4 = null;
        Row row5 = null;


        // 제목 컬럼 스타일
        XSSFCellStyle titleBgc = (XSSFCellStyle) workbook.createCellStyle();//0번째열 스타일
//        titleBgc.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex()); // 배경색
        XSSFColor mColor = new XSSFColor(new Color(224,234,246), null);
        titleBgc.setFillForegroundColor(mColor); // 배경색
        titleBgc.setFillPattern(FillPatternType.SOLID_FOREGROUND);// 채우기

        // 워크시트 생성
        Sheet sheet = workbook.createSheet("sheet1");

        //sheet 컬럼 폭 수정
        sheet.setColumnWidth(0,3800);
        sheet.setColumnWidth(1,2400);
        sheet.setColumnWidth(2,4600);
        sheet.setColumnWidth(3,6600);
        sheet.setColumnWidth(4,4800);
        sheet.setColumnWidth(5,6600);
        sheet.setColumnWidth(6,9200);
        sheet.setColumnWidth(7,2400);
        sheet.setColumnWidth(8,2400);

        row0 = sheet.createRow(0);


        row1 = sheet.createRow(1);

        row0.createCell(0).setCellValue("출고희망일");
        row0.createCell(1).setCellValue("출고센터");
        row0.createCell(2).setCellValue("출고형태");
        row0.createCell(3).setCellValue("도착지주소");
        row0.createCell(4).setCellValue("도착지 담당자 정보1");
        row0.createCell(5).setCellValue("도착지 담당자 정보2");
        row0.createCell(6).setCellValue("도착지 담당자 정보3");

        for(int i=0; i<7; i++){
            cell = row0.getCell(i);
            cell.setCellStyle(titleBgc);
        }

        row1.createCell(0).setCellValue("");
        row1.createCell(1).setCellValue("");
        row1.createCell(2).setCellValue("");
        row1.createCell(3).setCellValue("");
        row1.createCell(4).setCellValue("");
        row1.createCell(5).setCellValue("");
        row1.createCell(6).setCellValue("");

        row2 = sheet.createRow(2);

        row2.createCell(0).setCellValue("출고방식");
        row2.createCell(1).setCellValue("젹재방식");
        row2.createCell(2).setCellValue("배송도착희망시간");
        row2.createCell(3).setCellValue("하차방법");
        row2.createCell(4).setCellValue("요청사항");

        for(int i=0; i<5; i++){
            cell = row2.getCell(i);
            cell.setCellStyle(titleBgc);
        }

        row3 = sheet.createRow(3);

        row3.createCell(0).setCellValue("3");
        row3.createCell(1).setCellValue("3");
        row3.createCell(2).setCellValue("3");
        row3.createCell(3).setCellValue("3");
        row3.createCell(4).setCellValue("3");


        row4 = sheet.createRow(4);

        row4.createCell(0).setCellValue("No.");
        row4.createCell(1).setCellValue("품목명");
        row4.createCell(2).setCellValue("품목코드");
        row4.createCell(3).setCellValue("보관/패킹방식");
        row4.createCell(4).setCellValue("입고회차");
        row4.createCell(5).setCellValue("출고가능수량");
        row4.createCell(6).setCellValue("입고일");
        row4.createCell(7).setCellValue("유통기한");
        row4.createCell(8).setCellValue("요청수량");

        for(int i=0; i<9; i++){
            cell = row4.getCell(i);
            cell.setCellStyle(titleBgc);
        }

        row5 = sheet.createRow(5);

        row5.createCell(0).setCellValue("5");
        row5.createCell(1).setCellValue("5");
        row5.createCell(2).setCellValue("5");
        row5.createCell(3).setCellValue("5");
        row5.createCell(4).setCellValue("5");
        row5.createCell(5).setCellValue("5");
        row5.createCell(6).setCellValue("5");
        row5.createCell(7).setCellValue("5");
        row5.createCell(8).setCellValue("5");


        response.setContentType("ms-vnd/excel");
        response.setHeader("Content-Disposition", "attachment;filename=studentList.xls");

        workbook.write(response.getOutputStream());
        workbook.close();

    }
}
