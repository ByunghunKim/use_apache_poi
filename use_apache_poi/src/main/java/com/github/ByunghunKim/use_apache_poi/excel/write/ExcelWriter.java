package com.github.ByunghunKim.use_apache_poi.excel.write;

import org.apache.commons.codec.binary.Hex;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

@Controller
public class ExcelWriter {

    @RequestMapping(value = "/")
    public String index() {
        return "index";
    }
    
    @RequestMapping(value = "/hello")
    public String hello() {
        return "index";
    }

    @RequestMapping(value = "/getForm")
    public String xlsxWiter() {
        System.out.println("보조기기 급여기 지급청구서 엑셀 파일 다운로드");

        // 워크북 생성
        XSSFWorkbook workbook = new XSSFWorkbook();

        // 가로/세로 중앙정렬
        CellStyle centerAlignStyle = workbook.createCellStyle();
        centerAlignStyle.setAlignment(HorizontalAlignment.CENTER);
        centerAlignStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // 문서 제목 폰트 스타일
        Font titleFont = workbook.createFont();
        titleFont.setFontHeight((short)320);    // 16pt
        titleFont.setBold(true);

        // 문서 제목 스타일 정의
        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        titleStyle.setBorderTop(BorderStyle.NONE);
        titleStyle.setBorderBottom(BorderStyle.NONE);
        titleStyle.setBorderLeft(BorderStyle.NONE);
        titleStyle.setBorderRight(BorderStyle.NONE);
        titleStyle.setFont(titleFont);

        // 문서 제목 아래 안내 문구. (색상이 어두운 난은 신청인이 적지 않으며...)
        Font guideFont = workbook.createFont();
        guideFont.setFontHeight((short)160);  // 8pt
        CellStyle guideStyle = workbook.createCellStyle();
        guideStyle.setBorderTop(BorderStyle.NONE);
        guideStyle.setBorderBottom(BorderStyle.NONE);
        guideStyle.setBorderLeft(BorderStyle.NONE);
        guideStyle.setBorderRight(BorderStyle.NONE);
        guideStyle.setFont(guideFont);


        // 접수번호 회색배경 영역 - 왼쪽테두리 없음
        Font grayFont = workbook.createFont();
        grayFont.setFontHeight((short)180);  // 9pt
        CellStyle grayNoneLeftStyle = workbook.createCellStyle();
        grayNoneLeftStyle.setBorderTop(BorderStyle.HAIR);
        grayNoneLeftStyle.setBorderBottom(BorderStyle.HAIR);
        grayNoneLeftStyle.setBorderLeft(BorderStyle.NONE);
        grayNoneLeftStyle.setBorderRight(BorderStyle.NONE);
        grayNoneLeftStyle.setTopBorderColor(IndexedColors.GREY_50_PERCENT.index);
        grayNoneLeftStyle.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.index);
        grayNoneLeftStyle.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.index);
        grayNoneLeftStyle.setRightBorderColor(IndexedColors.GREY_50_PERCENT.index);
        grayNoneLeftStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
        grayNoneLeftStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        grayNoneLeftStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        grayNoneLeftStyle.setFont(grayFont);


        // 접수번호 회색배경 영역 - 왼쪽테두리 없음 + 중앙정렬
        CellStyle grayNoneLeftAlignCenterStyle = workbook.createCellStyle();
        grayNoneLeftAlignCenterStyle.setBorderTop(BorderStyle.HAIR);
        grayNoneLeftAlignCenterStyle.setBorderBottom(BorderStyle.HAIR);
        grayNoneLeftAlignCenterStyle.setBorderLeft(BorderStyle.NONE);
        grayNoneLeftAlignCenterStyle.setBorderRight(BorderStyle.HAIR);
        grayNoneLeftAlignCenterStyle.setTopBorderColor(IndexedColors.GREY_50_PERCENT.index);
        grayNoneLeftAlignCenterStyle.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.index);
        grayNoneLeftAlignCenterStyle.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.index);
        grayNoneLeftAlignCenterStyle.setRightBorderColor(IndexedColors.GREY_50_PERCENT.index);
        grayNoneLeftAlignCenterStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
        grayNoneLeftAlignCenterStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        grayNoneLeftAlignCenterStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        grayNoneLeftAlignCenterStyle.setAlignment(HorizontalAlignment.CENTER);
        grayNoneLeftAlignCenterStyle.setWrapText(true);
        grayNoneLeftAlignCenterStyle.setFont(grayFont);

        // 접수번호 회색배경 영역 - 왼쪽테두리 있음
        Font grayFont2 = workbook.createFont();
        grayFont2.setFontHeight((short)180);  // 9pt
        CellStyle grayThinLeftStyle = workbook.createCellStyle();
        grayThinLeftStyle.setBorderTop(BorderStyle.HAIR);
        grayThinLeftStyle.setBorderBottom(BorderStyle.HAIR);
        grayThinLeftStyle.setBorderLeft(BorderStyle.HAIR);
        grayThinLeftStyle.setBorderRight(BorderStyle.NONE);
        grayThinLeftStyle.setTopBorderColor(IndexedColors.GREY_50_PERCENT.index);
        grayThinLeftStyle.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.index);
        grayThinLeftStyle.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.index);
        grayThinLeftStyle.setRightBorderColor(IndexedColors.GREY_50_PERCENT.index);
        grayThinLeftStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
        grayThinLeftStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        grayThinLeftStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        grayThinLeftStyle.setFont(grayFont2);


        // 급여를 받을 사람 제목 영역
        Font table1TitleFont = workbook.createFont();
        table1TitleFont.setFontHeight((short)240);  // 12pt
        CellStyle table1TitleStyle = workbook.createCellStyle();
        table1TitleStyle.setBorderTop(BorderStyle.HAIR);
        table1TitleStyle.setBorderBottom(BorderStyle.HAIR);
        table1TitleStyle.setBorderLeft(BorderStyle.NONE);
        table1TitleStyle.setBorderRight(BorderStyle.NONE);
        table1TitleStyle.setTopBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table1TitleStyle.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table1TitleStyle.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table1TitleStyle.setRightBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table1TitleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        table1TitleStyle.setFont(table1TitleFont);
        table1TitleStyle.setWrapText(true);

        // 급여를 받을 사람 항목 영역
        Font table1Font = workbook.createFont();
        table1Font.setFontHeight((short)160);   // 8pt
        CellStyle table1Style = workbook.createCellStyle();
        table1Style.setBorderTop(BorderStyle.HAIR);
        table1Style.setBorderBottom(BorderStyle.HAIR);
        table1Style.setBorderLeft(BorderStyle.HAIR);
        table1Style.setBorderRight(BorderStyle.NONE);
        table1Style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table1Style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table1Style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table1Style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table1Style.setVerticalAlignment(VerticalAlignment.TOP);
        table1Style.setFont(table1Font);
        table1Style.setWrapText(true);

        // 보조기기 영역 - 코드
        Font table2CodeGrayFont = workbook.createFont();
        table2CodeGrayFont.setFontHeight((short)160);  // 8pt
        CellStyle table2CodeGrayStyle = workbook.createCellStyle();
        table2CodeGrayStyle.setBorderTop(BorderStyle.HAIR);
        table2CodeGrayStyle.setBorderBottom(BorderStyle.HAIR);
        table2CodeGrayStyle.setBorderLeft(BorderStyle.HAIR);
        table2CodeGrayStyle.setBorderRight(BorderStyle.NONE);
        table2CodeGrayStyle.setTopBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table2CodeGrayStyle.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table2CodeGrayStyle.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table2CodeGrayStyle.setRightBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table2CodeGrayStyle.setVerticalAlignment(VerticalAlignment.TOP);
        table2CodeGrayStyle.setAlignment(HorizontalAlignment.LEFT);
        table2CodeGrayStyle.setFont(table2CodeGrayFont);
        table2CodeGrayStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
        table2CodeGrayStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);


        // 기준액 제목 항목 영역
        Font table2TitleFont = workbook.createFont();
        table2TitleFont.setFontHeight((short)150);   // 7pt
        CellStyle table2TitleStyle = workbook.createCellStyle();
        table2TitleStyle.setBorderTop(BorderStyle.HAIR);
        table2TitleStyle.setBorderBottom(BorderStyle.HAIR);
        table2TitleStyle.setBorderLeft(BorderStyle.HAIR);
        table2TitleStyle.setBorderRight(BorderStyle.NONE);
        table2TitleStyle.setTopBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table2TitleStyle.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table2TitleStyle.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table2TitleStyle.setRightBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table2TitleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        table2TitleStyle.setAlignment(HorizontalAlignment.CENTER);
        table2TitleStyle.setFont(table2TitleFont);
        table2TitleStyle.setWrapText(true);

        Font table2TitleFont2 = workbook.createFont();
        table2TitleFont2.setFontHeight((short)160);   // 8pt
        CellStyle table2TitleStyle2 = workbook.createCellStyle();
        table2TitleStyle2.setBorderTop(BorderStyle.HAIR);
        table2TitleStyle2.setBorderBottom(BorderStyle.HAIR);
        table2TitleStyle2.setBorderLeft(BorderStyle.HAIR);
        table2TitleStyle2.setBorderRight(BorderStyle.NONE);
        table2TitleStyle2.setTopBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table2TitleStyle2.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table2TitleStyle2.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table2TitleStyle2.setRightBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table2TitleStyle2.setVerticalAlignment(VerticalAlignment.CENTER);
        table2TitleStyle2.setAlignment(HorizontalAlignment.CENTER);
        table2TitleStyle2.setFont(table2TitleFont2);
        table2TitleStyle2.setWrapText(true);

        // 기준액 가격입력란(원) 항목 영역 - 왼쪽 선 없음
        Font table2NoneLeftFont = workbook.createFont();
        table2NoneLeftFont.setFontHeight((short)260);  // 13pt
        CellStyle table2NoneLeftStyle = workbook.createCellStyle();
        table2NoneLeftStyle.setBorderTop(BorderStyle.HAIR);
        table2NoneLeftStyle.setBorderBottom(BorderStyle.HAIR);
        table2NoneLeftStyle.setBorderLeft(BorderStyle.NONE);
        table2NoneLeftStyle.setBorderRight(BorderStyle.NONE);
        table2NoneLeftStyle.setTopBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table2NoneLeftStyle.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table2NoneLeftStyle.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table2NoneLeftStyle.setRightBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table2NoneLeftStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        table2NoneLeftStyle.setAlignment(HorizontalAlignment.RIGHT);
        table2NoneLeftStyle.setFont(table2NoneLeftFont);

        // 기준액 가격입력란(원) 항목 영역 - 왼쪽 선 있음
        Font table2ThinLeftFont = workbook.createFont();
        table2ThinLeftFont.setFontHeight((short)260);  // 13pt
        CellStyle table2ThinLeftStyle = workbook.createCellStyle();
        table2ThinLeftStyle.setBorderTop(BorderStyle.HAIR);
        table2ThinLeftStyle.setBorderBottom(BorderStyle.HAIR);
        table2ThinLeftStyle.setBorderLeft(BorderStyle.HAIR);
        table2ThinLeftStyle.setBorderRight(BorderStyle.NONE);
        table2ThinLeftStyle.setTopBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table2ThinLeftStyle.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table2ThinLeftStyle.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table2ThinLeftStyle.setRightBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table2ThinLeftStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        table2ThinLeftStyle.setAlignment(HorizontalAlignment.RIGHT);
        table2ThinLeftStyle.setFont(table2ThinLeftFont);



        // 10 수령계좌 영역 - 왼쪽 선 없음
        Font table6NoneLeftFont = workbook.createFont();
        table6NoneLeftFont.setFontHeight((short)180);  // 9pt
        CellStyle table6NoneLeftStyle = workbook.createCellStyle();
        table6NoneLeftStyle.setBorderTop(BorderStyle.HAIR);
        table6NoneLeftStyle.setBorderBottom(BorderStyle.HAIR);
        table6NoneLeftStyle.setBorderLeft(BorderStyle.NONE);
        table6NoneLeftStyle.setBorderRight(BorderStyle.NONE);
        table6NoneLeftStyle.setTopBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table6NoneLeftStyle.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table6NoneLeftStyle.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table6NoneLeftStyle.setRightBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table6NoneLeftStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        table6NoneLeftStyle.setAlignment(HorizontalAlignment.CENTER);
        table6NoneLeftStyle.setFont(table6NoneLeftFont);
        table6NoneLeftStyle.setWrapText(true);

        // 급여를 받을 사람 항목 영역 - 설명문
        Font table7Font = workbook.createFont();
        table7Font.setFontHeight((short)200);   // 10pt
        CellStyle table7Style = workbook.createCellStyle();
        table7Style.setBorderTop(BorderStyle.HAIR);
        table7Style.setBorderBottom(BorderStyle.NONE);
        table7Style.setBorderLeft(BorderStyle.NONE);
        table7Style.setBorderRight(BorderStyle.NONE);
        table7Style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table7Style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table7Style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table7Style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table7Style.setFont(table7Font);
        table7Style.setVerticalAlignment(VerticalAlignment.CENTER);
        table7Style.setAlignment(HorizontalAlignment.CENTER);

        // 급여를 받을 사람 항목 영역 - 년월일 청구인 등
        Font table7SignFont = workbook.createFont();
        table7SignFont.setFontHeight((short)200);   // 10pt
        CellStyle table7SignStyle = workbook.createCellStyle();
        table7SignStyle.setBorderTop(BorderStyle.NONE);
        table7SignStyle.setBorderBottom(BorderStyle.NONE);
        table7SignStyle.setBorderLeft(BorderStyle.NONE);
        table7SignStyle.setBorderRight(BorderStyle.NONE);
        table7SignStyle.setFont(table7SignFont);
        table7SignStyle.setVerticalAlignment(VerticalAlignment.TOP);
        table7SignStyle.setAlignment(HorizontalAlignment.RIGHT);

        // 급여를 받을 사람 항목 영역 - 국민건강보험공단이사장 귀하
        Font table7NoneLineFont = workbook.createFont();
        table7NoneLineFont.setBold(true);
        table7NoneLineFont.setFontHeight((short)240);   // 12pt
        CellStyle table7NoneLineStyle = workbook.createCellStyle();
        table7NoneLineStyle.setBorderTop(BorderStyle.HAIR);
        table7NoneLineStyle.setBorderBottom(BorderStyle.HAIR);
        table7NoneLineStyle.setBorderLeft(BorderStyle.NONE);
        table7NoneLineStyle.setBorderRight(BorderStyle.NONE);
        table7NoneLineStyle.setTopBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table7NoneLineStyle.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table7NoneLineStyle.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table7NoneLineStyle.setRightBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table7NoneLineStyle.setFont(table7NoneLineFont);
        table7NoneLineStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        table7NoneLineStyle.setAlignment(HorizontalAlignment.LEFT);

        // 정보이용동의서
        Font table8Font = workbook.createFont();
        table8Font.setFontHeight((short)240);   // 12pt
        table8Font.setBold(true);
        CellStyle table8Style = workbook.createCellStyle();
        table8Style.setBorderTop(BorderStyle.MEDIUM);
        table8Style.setBorderBottom(BorderStyle.NONE);
        table8Style.setBorderLeft(BorderStyle.NONE);
        table8Style.setBorderRight(BorderStyle.NONE);
        table8Style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table8Style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table8Style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table8Style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table8Style.setFont(table8Font);
        table8Style.setVerticalAlignment(VerticalAlignment.CENTER);
        table8Style.setAlignment(HorizontalAlignment.CENTER);
        table8Style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
        table8Style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        table8Style.setVerticalAlignment(VerticalAlignment.CENTER);

        // 정보이용동의서 설명문
        Font table8GuideFont = workbook.createFont();
        table8GuideFont.setFontHeight((short)200);   // 10pt
        CellStyle table8GuideStyle = workbook.createCellStyle();
        table8GuideStyle.setBorderTop(BorderStyle.NONE);
        table8GuideStyle.setBorderBottom(BorderStyle.NONE);
        table8GuideStyle.setBorderLeft(BorderStyle.NONE);
        table8GuideStyle.setBorderRight(BorderStyle.NONE);
        table8GuideStyle.setFont(table8GuideFont);
        table8GuideStyle.setVerticalAlignment(VerticalAlignment.TOP);
        table8GuideStyle.setAlignment(HorizontalAlignment.LEFT);
        table8GuideStyle.setWrapText(true);

        // 하단 굵은선
        Font table8BottomFont = workbook.createFont();
        table8BottomFont.setFontHeight((short)240);   // 12pt
        CellStyle table8MiddleBottomStyle = workbook.createCellStyle();
        table8MiddleBottomStyle.setBorderTop(BorderStyle.NONE);
        table8MiddleBottomStyle.setBorderBottom(BorderStyle.MEDIUM);
        table8MiddleBottomStyle.setBorderLeft(BorderStyle.NONE);
        table8MiddleBottomStyle.setBorderRight(BorderStyle.NONE);
        table8MiddleBottomStyle.setTopBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table8MiddleBottomStyle.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table8MiddleBottomStyle.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table8MiddleBottomStyle.setRightBorderColor(IndexedColors.GREY_50_PERCENT.index);
        table8MiddleBottomStyle.setFont(table8BottomFont);
        table8MiddleBottomStyle.setVerticalAlignment(VerticalAlignment.TOP);
        table8MiddleBottomStyle.setAlignment(HorizontalAlignment.RIGHT);



// 1 page start
        // 워크시트 생성
        XSSFSheet sheet = workbook.createSheet("보조기기 급여비 지급청구서");
        sheet.setMargin(HSSFSheet.TopMargin, 0.65);
        sheet.setMargin(HSSFSheet.BottomMargin, 0.65);
        sheet.setMargin(HSSFSheet.LeftMargin, 0.65);
        sheet.setMargin(HSSFSheet.RightMargin, 0.65);
        // 행 생성
        XSSFRow row = null;
        // 셀 생성
        XSSFCell cell;


        // 제목 정보 구성
        row = sheet.createRow(0);
        row.setHeight((short) 600);
        cell = row.createCell(0);
        for(int i=0; i<34; i++) {
            // 컬럼 폭
            sheet.setColumnWidth(i, 620);
        }
        // 제목 영역
        cell = row.createCell(0);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("보조기기 급여비 지급청구서");
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 33));


        // 문서 제목 아래 안내 문구 영역(색상이 어두운 난은...)
        row = sheet.createRow(1);
        row.setHeight((short) 285); // 0.50mm
        cell = row.createCell(0);
        cell.setCellStyle(guideStyle);
        cell.setCellValue("※ 색상이 어두운 난은 신청인이 적지 않으며, [    ]에는 해당되는 곳에 √ 표시를 합니다.");
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 31));
        cell = row.createCell(32);
        cell.setCellStyle(guideStyle);
        cell.setCellValue("(앞쪽)");
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 32, 33));


        // 접수번호
        row = sheet.createRow(2);
        row.setHeight((short) 390); //0.69mm
        // 셀 테두리 스타일 적용 위해서 각 셀을 만들어주어야 한다.
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(grayNoneLeftStyle);
            cell.setCellValue("접수번호");
        }
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 11));
        cell = row.createCell(12);
        cell.setCellStyle(grayThinLeftStyle);
        cell.setCellValue("접수일");
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 12, 23));
        cell = row.createCell(24);
        cell.setCellStyle(grayThinLeftStyle);
        cell.setCellValue("처리기간          7일");
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 24, 33));


        // 본인부담액
        row = sheet.createRow(3);
        row.setHeight((short) 390); //0.69mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(grayNoneLeftAlignCenterStyle);
            cell.setCellValue("본인부담액\r\n경감 대상자");
        }
        sheet.addMergedRegion(new CellRangeAddress(3, 4, 0, 5));
        for(int i=6; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(grayThinLeftStyle);
            cell.setCellValue(" [      ] 「국민건강보험법 시행령」  별표 2 제3호라목1)에 해당하는 사람");
        }
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 6, 33));

        row = sheet.createRow(4);
        row.setHeight((short) 390); //0.69mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(grayNoneLeftAlignCenterStyle);
        }
        for(int i=6; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(grayThinLeftStyle);
            cell.setCellValue(" [      ] 「국민건강보험법 시행령」  별표 2 제3호라목2)에 해당하는 사람");
        }
        sheet.addMergedRegion(new CellRangeAddress(4, 4, 6, 33));



        // 공백 행
        row = sheet.createRow(5);
        row.setHeight((short) 110);
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 0, 33));

        // 급여를 받을 사람
        row = sheet.createRow(6);
        row.setHeight((short) 448); //0.79mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table1TitleStyle);
            cell.setCellValue("① 급여를\r\n      받을 사람");
        }
        sheet.addMergedRegion(new CellRangeAddress(6, 8, 0, 5));
        cell = row.createCell(6);
        cell.setCellStyle(table1Style);
        cell.setCellValue("성명");
        sheet.addMergedRegion(new CellRangeAddress(6, 6, 6, 15));
        cell = row.createCell(16);
        cell.setCellStyle(table1Style);
        cell.setCellValue("주민(외국인)등록번호");
        sheet.addMergedRegion(new CellRangeAddress(6, 6, 16, 33));

        row = sheet.createRow(7);
        row.setHeight((short) 448); //0.79mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table1TitleStyle);
            cell = row.createCell(6);
            cell.setCellStyle(table1Style);
            cell.setCellValue("집 전화번호");
        }
        sheet.addMergedRegion(new CellRangeAddress(7, 7, 6, 15));
        cell = row.createCell(16);
        cell.setCellStyle(table1Style);
        cell.setCellValue("휴대전화번호");
        sheet.addMergedRegion(new CellRangeAddress(7, 7, 16, 33));

        row = sheet.createRow(8);
        row.setHeight((short) 448); //0.79mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table1TitleStyle);
            cell = row.createCell(6);
            cell.setCellStyle(table1Style);
            cell.setCellValue("장애명");
        }
        sheet.addMergedRegion(new CellRangeAddress(8, 8, 6, 15));
        cell = row.createCell(16);
        cell.setCellStyle(table1Style);
        cell.setCellValue("장애 정도             [       ] 심한 장애                   [       ] 심하지 않은 장애");
        sheet.addMergedRegion(new CellRangeAddress(8, 8, 16, 33));

        // 공백 행
        row = sheet.createRow(9);
        row.setHeight((short) 110);
        sheet.addMergedRegion(new CellRangeAddress(9, 9, 0, 33));

        // 보조기기
        row = sheet.createRow(10);
        row.setHeight((short) 448); //0.79mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table1TitleStyle);
            cell.setCellValue("② 보조기기");
        }
        sheet.addMergedRegion(new CellRangeAddress(10, 11, 0, 5));
        cell = row.createCell(6);
        cell.setCellStyle(table1Style);
        cell.setCellValue("명칭");
        sheet.addMergedRegion(new CellRangeAddress(10, 10, 6, 15));
        cell = row.createCell(16);
        cell.setCellStyle(table2CodeGrayStyle);
        cell.setCellValue("코드");
        sheet.addMergedRegion(new CellRangeAddress(10, 10, 16, 33));

        row = sheet.createRow(11);
        row.setHeight((short) 448); //0.79mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table1TitleStyle);
            cell = row.createCell(6);
            cell.setCellStyle(table1Style);
            cell.setCellValue("구입일");
        }
        sheet.addMergedRegion(new CellRangeAddress(11, 11, 6, 33));

        // 공백 행
        row = sheet.createRow(12);
        row.setHeight((short) 110);
        sheet.addMergedRegion(new CellRangeAddress(12, 12, 0, 33));

        // 제품정보
        row = sheet.createRow(13);
        row.setHeight((short) 448); //0.79mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table1TitleStyle);
            cell.setCellValue("③ 제품정보");
        }
        sheet.addMergedRegion(new CellRangeAddress(13, 14, 0, 5));
        cell = row.createCell(6);
        cell.setCellStyle(table1Style);
        cell.setCellValue("모델명");
        sheet.addMergedRegion(new CellRangeAddress(13, 13, 6, 13));
        cell = row.createCell(14);
        cell.setCellStyle(table1Style);
        cell.setCellValue("제조(수입)업소명");
        sheet.addMergedRegion(new CellRangeAddress(13, 13, 14, 23));
        cell = row.createCell(24);
        cell.setCellStyle(table1Style);
        cell.setCellValue("제조일");
        sheet.addMergedRegion(new CellRangeAddress(13, 13, 24, 33));

        row = sheet.createRow(14);
        row.setHeight((short) 448); //0.79mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table1TitleStyle);
            cell = row.createCell(6);
            cell.setCellStyle(table1Style);
            cell.setCellValue("제품제조번호");
        }
        sheet.addMergedRegion(new CellRangeAddress(14, 14, 6, 15));
        cell = row.createCell(16);
        cell.setCellStyle(table1Style);
        cell.setCellValue("표준코드");
        sheet.addMergedRegion(new CellRangeAddress(14, 14, 16, 33));

        // 공백 행
        row = sheet.createRow(15);
        row.setHeight((short) 110);
        sheet.addMergedRegion(new CellRangeAddress(15, 15, 0, 33));

        // 구입처
        row = sheet.createRow(16);
        row.setHeight((short) 448); //0.79mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table1TitleStyle);
            cell.setCellValue("④ 구입처");
        }
        sheet.addMergedRegion(new CellRangeAddress(16, 18, 0, 5));
        cell = row.createCell(6);
        cell.setCellStyle(table1Style);
        cell.setCellValue("명칭");
        sheet.addMergedRegion(new CellRangeAddress(16, 16, 6, 15));
        cell = row.createCell(16);
        cell.setCellStyle(table1Style);
        cell.setCellValue("대표자");
        sheet.addMergedRegion(new CellRangeAddress(16, 16, 16, 33));

        row = sheet.createRow(17);
        row.setHeight((short) 448); //0.79mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table1TitleStyle);
            cell = row.createCell(6);
            cell.setCellStyle(table1Style);
            cell.setCellValue("사업자등록번호");
        }
        sheet.addMergedRegion(new CellRangeAddress(17, 17, 6, 15));
        cell = row.createCell(16);
        cell.setCellStyle(table1Style);
        cell.setCellValue("전화번호");
        sheet.addMergedRegion(new CellRangeAddress(17, 17, 16, 33));

        row = sheet.createRow(18);
        row.setHeight((short) 448); //0.79mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table1TitleStyle);
            cell = row.createCell(6);
            cell.setCellStyle(table1Style);
            cell.setCellValue("주소(미등록 업소만 기록합니다)");
        }
        sheet.addMergedRegion(new CellRangeAddress(18, 18, 6, 33));

        // 공백 행
        row = sheet.createRow(19);
        row.setHeight((short) 110);
        sheet.addMergedRegion(new CellRangeAddress(19, 19, 0, 33));

        // 기준액
        row = sheet.createRow(20);
        row.setHeight((short) 795); //1.40mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table1TitleStyle);
            cell.setCellValue("⑤ 기준액");
        }
        sheet.addMergedRegion(new CellRangeAddress(20, 20, 0, 5));
        cell = row.createCell(6);
        cell.setCellStyle(table2TitleStyle);
        cell.setCellValue("⑥ 고시금액\n(전동휠체어, 전동스쿠터 및\n자세한 보조용구만 적습니다)");
        sheet.addMergedRegion(new CellRangeAddress(20, 20, 6, 12));
        cell = row.createCell(13);
        cell.setCellStyle(table2TitleStyle2);
        cell.setCellValue("⑦ 실구입금액\n(⑧+⑨)");
        sheet.addMergedRegion(new CellRangeAddress(20, 20, 13, 19));
        cell = row.createCell(20);
        cell.setCellStyle(table2TitleStyle2);
        cell.setCellValue("⑧ 본인부담액");
        sheet.addMergedRegion(new CellRangeAddress(20, 20, 20, 26));
        cell = row.createCell(27);
        cell.setCellStyle(table2TitleStyle2);
        cell.setCellValue("⑨ 청구금액");
        sheet.addMergedRegion(new CellRangeAddress(20, 20, 27, 33));
        //가격란1
        row = sheet.createRow(21);
        row.setHeight((short) 448); //0.79mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table2NoneLeftStyle);
            cell.setCellValue("원");
        }
        sheet.addMergedRegion(new CellRangeAddress(21, 21, 0, 5));
        cell = row.createCell(6);
        cell.setCellStyle(table2ThinLeftStyle);
        cell.setCellValue("원");
        sheet.addMergedRegion(new CellRangeAddress(21, 21, 6, 12));
        cell = row.createCell(13);
        cell.setCellStyle(table2ThinLeftStyle);
        cell.setCellValue("원");
        sheet.addMergedRegion(new CellRangeAddress(21, 21, 13, 19));
        cell = row.createCell(20);
        cell.setCellStyle(table2ThinLeftStyle);
        cell.setCellValue("원");
        sheet.addMergedRegion(new CellRangeAddress(21, 21, 20, 26));
        cell = row.createCell(27);
        cell.setCellStyle(table2ThinLeftStyle);
        cell.setCellValue("원");
        sheet.addMergedRegion(new CellRangeAddress(21, 21, 27, 33));
        //가격란2
        row = sheet.createRow(22);
        row.setHeight((short) 448); //0.79mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table2NoneLeftStyle);
            cell.setCellValue("원");
        }
        sheet.addMergedRegion(new CellRangeAddress(22, 22, 0, 5));
        cell = row.createCell(6);
        cell.setCellStyle(table2ThinLeftStyle);
        cell.setCellValue("원");
        sheet.addMergedRegion(new CellRangeAddress(22, 22, 6, 12));
        cell = row.createCell(13);
        cell.setCellStyle(table2ThinLeftStyle);
        cell.setCellValue("원");
        sheet.addMergedRegion(new CellRangeAddress(22, 22, 13, 19));
        cell = row.createCell(20);
        cell.setCellStyle(table2ThinLeftStyle);
        cell.setCellValue("원");
        sheet.addMergedRegion(new CellRangeAddress(22, 22, 20, 26));
        cell = row.createCell(27);
        cell.setCellStyle(table2ThinLeftStyle);
        cell.setCellValue("원");
        sheet.addMergedRegion(new CellRangeAddress(22, 22, 27, 33));
        //가격란3
        row = sheet.createRow(23);
        row.setHeight((short) 448); //0.79mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table2NoneLeftStyle);
            cell.setCellValue("원");
        }
        sheet.addMergedRegion(new CellRangeAddress(23, 23, 0, 5));
        cell = row.createCell(6);
        cell.setCellStyle(table2ThinLeftStyle);
        cell.setCellValue("원");
        sheet.addMergedRegion(new CellRangeAddress(23, 23, 6, 12));
        cell = row.createCell(13);
        cell.setCellStyle(table2ThinLeftStyle);
        cell.setCellValue("원");
        sheet.addMergedRegion(new CellRangeAddress(23, 23, 13, 19));
        cell = row.createCell(20);
        cell.setCellStyle(table2ThinLeftStyle);
        cell.setCellValue("원");
        sheet.addMergedRegion(new CellRangeAddress(23, 23, 20, 26));
        cell = row.createCell(27);
        cell.setCellStyle(table2ThinLeftStyle);
        cell.setCellValue("원");
        sheet.addMergedRegion(new CellRangeAddress(23, 23, 27, 33));
        //가격란4
        row = sheet.createRow(24);
        row.setHeight((short) 448); //0.79mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table2NoneLeftStyle);
            cell.setCellValue("원");
        }
        sheet.addMergedRegion(new CellRangeAddress(24, 24, 0, 5));
        cell = row.createCell(6);
        cell.setCellStyle(table2ThinLeftStyle);
        cell.setCellValue("원");
        sheet.addMergedRegion(new CellRangeAddress(24, 24, 6, 12));
        cell = row.createCell(13);
        cell.setCellStyle(table2ThinLeftStyle);
        cell.setCellValue("원");
        sheet.addMergedRegion(new CellRangeAddress(24, 24, 13, 19));
        cell = row.createCell(20);
        cell.setCellStyle(table2ThinLeftStyle);
        cell.setCellValue("원");
        sheet.addMergedRegion(new CellRangeAddress(24, 24, 20, 26));
        cell = row.createCell(27);
        cell.setCellStyle(table2ThinLeftStyle);
        cell.setCellValue("원");
        sheet.addMergedRegion(new CellRangeAddress(24, 24, 27, 33));


        // 공백 행
        row = sheet.createRow(25);
        row.setHeight((short) 110);
        sheet.addMergedRegion(new CellRangeAddress(25, 25, 0, 33));

        // 수령계좌
        row = sheet.createRow(26);
        row.setHeight((short) 285); // 0.50mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table6NoneLeftStyle);
            cell.setCellValue("⑩\r\n수령\r\n계좌");
        }
        sheet.addMergedRegion(new CellRangeAddress(26, 28, 0, 1));
        cell = row.createCell(2);
        cell.setCellStyle(table1Style);
        cell.setCellValue("[      ] 가입자 또는 피부양자 계좌");
        sheet.addMergedRegion(new CellRangeAddress(26, 26, 2, 12));
        cell = row.createCell(13);
        cell.setCellStyle(table1Style);
        cell.setCellValue("금융기관명");
        sheet.addMergedRegion(new CellRangeAddress(26, 27, 13, 20));
        cell = row.createCell(21);
        cell.setCellStyle(table1Style);
        cell.setCellValue("계좌번호");
        sheet.addMergedRegion(new CellRangeAddress(26, 27, 21, 33));

        row = sheet.createRow(27);
        row.setHeight((short) 285); // 0.50mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table1TitleStyle);
            cell = row.createCell(2);
            cell.setCellStyle(table1Style);
            cell.setCellValue("[      ] 보조기기 제조∙대여∙판매업소 계좌");
        }
        sheet.addMergedRegion(new CellRangeAddress(27, 27, 2, 12));
        cell = row.createCell(13);
        cell.setCellStyle(table1Style);
        cell = row.createCell(21);
        cell.setCellStyle(table1Style);

        row = sheet.createRow(28);
        row.setHeight((short) 570); //1.00mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table6NoneLeftStyle);
            cell = row.createCell(2);
            cell.setCellStyle(table1Style);
            cell.setCellValue("[      ] 진료받은 사람 본인의 요양비 등\r\n         수급계좌(압류방지 계좌)");
        }
        sheet.addMergedRegion(new CellRangeAddress(28, 28, 2, 12));
        cell = row.createCell(13);
        cell.setCellStyle(table1Style);
        cell.setCellValue("예금주");
        sheet.addMergedRegion(new CellRangeAddress(28, 28, 13, 20));
        cell = row.createCell(21);
        cell.setCellStyle(table1Style);
        cell.setCellValue("주민(외국인)등록번호 또는 사업자등록번호");
        sheet.addMergedRegion(new CellRangeAddress(28, 28, 21, 33));

        // 공백 행
        row = sheet.createRow(29);
        row.setHeight((short) 110);
        sheet.addMergedRegion(new CellRangeAddress(29, 29, 0, 33));

        // 국민건강보험법 시행규칙
        row = sheet.createRow(30);
        row.setHeight((short) 448); //0.79mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table7Style);
            cell.setCellValue("「국민건강보험법 시행규칙」 제26조제2항∙제4항에 따라 위와 같이 보조기기 급여비의 지급을 청구합니다.");
        }
        sheet.addMergedRegion(new CellRangeAddress(30, 30, 0, 33));

        // 년월일
        row = sheet.createRow(31);
        row.setHeight((short) 420); // 0.58mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table7SignStyle);
            cell.setCellValue("년                월                일");
        }
        sheet.addMergedRegion(new CellRangeAddress(31, 31, 0, 33));

        // 청구인
        row = sheet.createRow(32);
        row.setHeight((short) 420); // 0.58mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table7SignStyle);
            cell = row.createCell(7);
            cell.setCellStyle(table7SignStyle);
            cell.setCellValue("⑪ 청구인");
        }
        sheet.addMergedRegion(new CellRangeAddress(32, 32, 7, 11));

        // 서명
        cell = row.createCell(17);
        cell.setCellStyle(table7SignStyle);
        cell.setCellValue("(서명 또는 인)  주민등록번호");
        sheet.addMergedRegion(new CellRangeAddress(32, 32, 17, 26));

        // 주민등록번호
        //sheet.addMergedRegion(new CellRangeAddress(32, 32, 23, 33));
        //cell = row.createCell(23);
        //cell.setCellStyle(table7SignStyle);
        //cell.setCellValue("");

        // 급여를 받을 사람과의 관계
        row = sheet.createRow(33);
        row.setHeight((short) 420); // 0.58mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table7SignStyle);
            cell = row.createCell(3);
            cell.setCellStyle(table7SignStyle);
            cell.setCellValue("급여를 받을 사람과의 관계");
        }
        sheet.addMergedRegion(new CellRangeAddress(33, 33, 3, 11));

        // 전화번호
        cell = row.createCell(17);
        cell.setCellStyle(table7SignStyle);
        cell.setCellValue("(휴대)전화번호");
        sheet.addMergedRegion(new CellRangeAddress(33, 33, 17, 26));


        // 국민건강보험공단 이사장
        row = sheet.createRow(34);
        row.setHeight((short) 420); // 0.58mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table7NoneLineStyle);
            cell.setCellValue("    국민건강보험공단 이사장 귀하");
        }
        sheet.addMergedRegion(new CellRangeAddress(34, 34, 0, 33));

        // 정보이용동의서
        row = sheet.createRow(35);
        row.setHeight((short) 330); // 0.58mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table8Style);
            cell.setCellValue("정보 이용 동의서");
        }
        sheet.addMergedRegion(new CellRangeAddress(35, 35, 0, 33));

        // 서명 동의
        row = sheet.createRow(36);
        row.setHeight((short) 680); // 1.20mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table8GuideStyle);
            cell.setCellValue("     본인은 위 보조기기 급여비의 지급 관련 정보(급여비 지급 여부∙품목, 사용 가능 기간 등)를 「사회보장기본법」 \r\n" +
                    "     제37조에 따라 사회보장정보시스템에 제공하는 것에 동의합니다.");
        }
        sheet.addMergedRegion(new CellRangeAddress(36, 36, 0, 33));

        // 급여를 받을 사람
        row = sheet.createRow(37);
        row.setHeight((short) 420); // 0.58mm
        for(int i=0; i<34; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(table8MiddleBottomStyle);
        }
        sheet.addMergedRegion(new CellRangeAddress(37, 37, 0, 8));
        cell = row.createCell(9);
        cell.setCellStyle(table8MiddleBottomStyle);
        cell.setCellValue("급여를 받을 사람");
        sheet.addMergedRegion(new CellRangeAddress(37, 37, 9, 16));
        cell = row.createCell(17);
        cell.setCellStyle(table8MiddleBottomStyle);
        sheet.addMergedRegion(new CellRangeAddress(37, 37, 17, 25));
        cell = row.createCell(26);
        cell.setCellStyle(table8MiddleBottomStyle);
        cell.setCellValue("(서명 또는 인)");
        sheet.addMergedRegion(new CellRangeAddress(37, 37, 26, 33));

// 1 page end
// 2 page start


        // 워크시트 생성
        XSSFSheet sheet2 = workbook.createSheet("유의사항_작성방법");
        sheet2.setMargin(HSSFSheet.TopMargin, 0.65);
        sheet2.setMargin(HSSFSheet.BottomMargin, 0.65);
        sheet2.setMargin(HSSFSheet.LeftMargin, 0.65);
        sheet2.setMargin(HSSFSheet.RightMargin, 0.65);
        // 행 생성
        XSSFRow row2 = null;
        // 셀 생성
        XSSFCell cell2;

        // 첨부서류 항목 영역
        Font docGuideFont = workbook.createFont();
        docGuideFont.setFontHeight((short)150);   // 7pt    // 8pt    (short)160
        CellStyle docGuideStyle = workbook.createCellStyle();
        docGuideStyle.setBorderTop(BorderStyle.MEDIUM);
        docGuideStyle.setBorderBottom(BorderStyle.HAIR);
        docGuideStyle.setBorderLeft(BorderStyle.NONE);
        docGuideStyle.setBorderRight(BorderStyle.NONE);
        docGuideStyle.setTopBorderColor(IndexedColors.GREY_50_PERCENT.index);
        docGuideStyle.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.index);
        docGuideStyle.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.index);
        docGuideStyle.setRightBorderColor(IndexedColors.GREY_50_PERCENT.index);
        docGuideStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        docGuideStyle.setAlignment(HorizontalAlignment.LEFT);
        docGuideStyle.setFont(docGuideFont);
        docGuideStyle.setWrapText(true);

        CellStyle docGuideStyle2 = workbook.createCellStyle();
        docGuideStyle2.setBorderTop(BorderStyle.MEDIUM);
        docGuideStyle2.setBorderBottom(BorderStyle.HAIR);
        docGuideStyle2.setBorderLeft(BorderStyle.HAIR);
        docGuideStyle2.setBorderRight(BorderStyle.NONE);
        docGuideStyle2.setTopBorderColor(IndexedColors.GREY_50_PERCENT.index);
        docGuideStyle2.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.index);
        docGuideStyle2.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.index);
        docGuideStyle2.setRightBorderColor(IndexedColors.GREY_50_PERCENT.index);
        docGuideStyle2.setVerticalAlignment(VerticalAlignment.CENTER);
        docGuideStyle2.setAlignment(HorizontalAlignment.LEFT);
        docGuideStyle2.setFont(docGuideFont);
        docGuideStyle2.setWrapText(true);

        // 첨부서류 항목 영역
        CellStyle docGuideStyle3 = workbook.createCellStyle();
        docGuideStyle3.setBorderTop(BorderStyle.NONE);
        docGuideStyle3.setBorderBottom(BorderStyle.HAIR);
        docGuideStyle3.setBorderLeft(BorderStyle.NONE);
        docGuideStyle3.setBorderRight(BorderStyle.NONE);
        docGuideStyle3.setTopBorderColor(IndexedColors.GREY_50_PERCENT.index);
        docGuideStyle3.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.index);
        docGuideStyle3.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.index);
        docGuideStyle3.setRightBorderColor(IndexedColors.GREY_50_PERCENT.index);
        docGuideStyle3.setVerticalAlignment(VerticalAlignment.CENTER);
        docGuideStyle3.setAlignment(HorizontalAlignment.LEFT);
        docGuideStyle3.setFont(docGuideFont);
        docGuideStyle3.setWrapText(true);

        

        // 첨부서류
        row2 = sheet2.createRow(0);
        row2.setHeight((short) 5200);
        cell2 = row2.createCell(0);
        for(int i=0; i<34; i++) {
            // 컬럼 폭
            sheet2.setColumnWidth(i, 620); 
        }
        // 첨부서류 설명란
        for(int i=0; i<3; i++) {
            cell2 = row2.createCell(i);
            cell2.setCellStyle(docGuideStyle);
            cell2.setCellValue("첨부서류");
        }
        sheet2.addMergedRegion(new CellRangeAddress(0, 0, 0, 2));

        for(int i=3; i<34; i++) {
            cell2 = row2.createCell(i);
            cell2.setCellStyle(docGuideStyle2);
            cell2.setCellValue(
                            "1. 활동형 수동휠체어, 틸팅형 수동휠체어, 리클라이닝형 수동휠체어, 전동휠체어, 전동스쿠터 및 이동식전동리프트에 대하여 보험급여\r\n" +
                            "     를 받으려는 경우: 별표7 제1호다목에 따라 공단에 등록한 보조기기 업소에서 발행한 세금계산서 1부 \r\n" +
                            "2. 자세보조용구에 대하여 보험급여를 받으려는 경우: 다음 각 목의 서류 \r\n" +
                            "     가. 「국민건강보험법 시행규칙」 별지 제23호서식의 보조기기 검수확인서 1부 \r\n" +
                            "     나. 별표7 제1호다목에 따라 공단에 등록한 보조기기 업소에서 발행한 세금계산서 1부 \r\n" +
                            "3. 그 밖의 보조기기에 대하여 보험급여를 받으려는 경우: 다음 각 목의 서류 \r\n" +
                            "     가. 「국민건강보험법 시행규칙」 별지 제22호서식 및 별지 제22호의2서식부터 제22호의4서식까지에 따른 보조기기 처방전과 해당 검 \r\n" +
                            "          사 결과 관련 서류 및 별지 제23호서식의 보조기기 검수확인서 각 1부. 다만, 지팡이목발, 흰 지팡이 및 보조기기의 소모품에 대하여 \r\n" +
                            "          보험급여를 받으려는 경우에는 보조기기 처방전과 해당 검사 결과 관련 서류 및 보조기기 검수확인서를 첨부하지 않으며, 일반형 \r\n" +
                            "          수동휠체어, 욕창예방방석, 욕창예방매트리스, 전·후방보행보조차, 돋보기 및 망원경에 대하여 보험급여를 받으려는 경우에는 보 \r\n" +
                            "          조기기 검수확인서를 첨부하지 않습니다. \r\n" +
                            "     나. 요양기관 또는 보조기기 제조·판매자가 발행한 세금계산서 1부 \r\n" +
                            "4. 보조기기 급여비를 보조기기의 제조·판매자에게 지급할 것을 신청한 경우에는 제1호부터 제3호까지의서류 중 해당하는 서류와 함께 \r\n" +
                            "     보조기기의 제조·판매자가 다음 각 목의 어느 하나에 해당하는 사람임을 증명하는 서류 1부를 제출합니다. 다만, 「국민건강 보험법 \r\n" +
                            "     시행규칙」 별표 7 제1호다목에 따라 공단에 등록한 보조기기 업소에서 구입한 경우, 지체장애 및 뇌병변장애에 대한 보행 보조를 위하\r\n" +
                            "     여 지팡이 또는 목발을 구입하였거나 시각장애에 대한 보행보조를 위하여 흰 지팡이를 구입한 경우 및 보조기기를 제조·수입한 업소 \r\n" +
                            "     에서 해당 보조기기의 소모품 중 전동휠체어 및 전동스쿠터용 전지를 구입한 경우에는 다음 각 목의 서류를 첨부하지 않습니다. \r\n" +
                            "     가. 「장애인복지법」에 따라 개설된 의지·보조기 제조·수리업자 \r\n" +
                            "     나. 「의료기기법」에 따라 허가받은 수입·제조·판매업자 \r\n" +
                            "     다. 「의료기기법」에 따라 신고한 전동휠체어·전동스쿠터용 전지(電池)의 수리업자 \r\n" +
                            "5. 전동휠체어, 전동스쿠터, 이동식전동리프트, 자세보조용구, 수동휠체어, 보청기, 욕창예방방석, 욕창예방매트리스, 전·후방보행 보\r\n" +
                            "     조차에 대하여 보험급여를 받으려는 경우에는 제1호부터 제4호까지의 서류 중 해당하는 서류와 함께 표준코드와 바코드를 확인 할 수\r\n" +
                            "     있는 보조기기 사진 1장을 제출합니다. \r\n" +
                            "6. 보조기기 급여비의 수령계좌가 진료받은 사람 본인의 요양비등 수급계좌(압류방지 계좌)인 경우에는 행복지킴이 통장 사본(계좌번호\r\n" +
                            "     가 기록되어 있는 면의 사본을 첨부합니다) 1부를 제출합니다.");
        }
        sheet2.addMergedRegion(new CellRangeAddress(0, 0, 3, 33));

        // 공백 행
        row2 = sheet2.createRow(1);
        row2.setHeight((short) 60);
        sheet2.addMergedRegion(new CellRangeAddress(1, 1, 0, 33));

        // 유의사항
        row2 = sheet2.createRow(2);
        row2.setHeight((short) 330); // 0.58mm
        for(int i=0; i<34; i++) {
            cell2 = row2.createCell(i);
            cell2.setCellStyle(table8Style);
            cell2.setCellValue("유의사항");
        }
        sheet2.addMergedRegion(new CellRangeAddress(2, 2, 0, 33));

        // 유의사항 본문
        row2 = sheet2.createRow(3);
        row2.setHeight((short) 900); // 1.20mm
        for(int i=0; i<34; i++) {
            cell2 = row2.createCell(i);
            cell2.setCellStyle(docGuideStyle3);
            cell2.setCellValue("※ 수동휠체어, 보청기, 전동휠체어, 전동스쿠터, 자세보조용구, 욕창예방방석, 욕창예방매트리스, 이동식전동리프트, 전·후방보행보조차의 경우에 \r\n" +
            "     는 공단에 등록된 품목에 대해서만 보험급여를 실시하며, 의지·보조기, 맞춤형 교정용 신발, 수동휠체어, 보청기, 전동휠체어, 전동스쿠터, 자세보 \r\n" +
            "     조용구, 욕창에방방석, 욕창예방매트리스, 이동식전동리프트, 전·후방보행보조차의 경우에는 공단에 등록된 업소에서 구입했을 때에만 보험급여 \r\n" +
            "     를 실시하므로 보조기기 구입 전 공단 등록 여부를 반드시 확인하시기 바랍니다.");
        }
        sheet2.addMergedRegion(new CellRangeAddress(3, 3, 0, 33));

        // 공백 행
        row2 = sheet2.createRow(4);
        row2.setHeight((short) 60);
        sheet2.addMergedRegion(new CellRangeAddress(4, 4, 0, 33));

        // 작성방법
        row2 = sheet2.createRow(5);
        row2.setHeight((short) 330); // 0.58mm
        for(int i=0; i<34; i++) {
            cell2 = row2.createCell(i);
            cell2.setCellStyle(table8Style);
            cell2.setCellValue("작성방법");
        }
        sheet2.addMergedRegion(new CellRangeAddress(5, 5, 0, 33));

        // 작성방법 본문
        row2 = sheet2.createRow(6);
        row2.setHeight((short) 7600); // 1.20mm
        for(int i=0; i<34; i++) {
            cell2 = row2.createCell(i);
            cell2.setCellStyle(docGuideStyle3);
            cell2.setCellValue("① 급여를 받을 사람: 급여를 받을 사람의 해당사항을 적습니다. \r\n" +
            "② 보조기기: 구입한 보조기기의 명칭과 보조기기를 구입한 연월일을 적습니다. \r\n" +
            "③ 제품정보: 구입한 보조기기의 모델명, 제조(수입)업소명, 제조 연월일, 제품제조번호 및 표준코드를 적습니다. \r\n" +
            "    ※ 제품제조번호와 표준코드는 전동휠체어, 전동스쿠터, 이동식전동리프트, 자세보조용구, 수동휠체어, 보청기, 욕창예방방석, 욕창예방매트리스,\r\n" +
            "        전·후방보행보조차를 구입한 경우에만 적습니다. \r\n" +
            "④ 구입처: 보조기기를 구입한 업소의 명칭, 대표자 성명, 사업자등록번호, 전화번호 및 주소(미등록 업소만 해당합니다)를 적습니다. \r\n" +
            "⑤ 기준액: 「국민건강보험법 시행규칙」 별표 7 제2호의 보조기기의 유형 및 구분 항목별 기준액을 적습니다. \r\n" +
            "       (예시: 자세보조용구 몸통 및 골반지지대/머리 및 목지지대를 동시에 장착한 경우에는 880,000원/210,000원을 각각 적습니다) \r\n" +
            "⑥ 고시금액: 전동휠체어, 전동스쿠터 및 자세보조용구만 보건복지부장관이 고시한 금액을 적습니다. \r\n" +
            "⑦ 실구입금액: 구입한 보조기기의 실제 구입금액(세금계산서의 금액을 말합니다)을 적고, 자세보조용구는 유형 및 구분 항목별로 실구입금액을 적 \r\n" +
            "       습니다. \r\n" +
            "⑧ 본인부담액: 실구입금액이 기준액보다 적은 경우에는 기준액·고시금액·실구입금액 중 가장 낮은 금액의 10%에 해당하는 금액을 적고, 실구입금 \r\n" +
            "       액이 기준액보다 많은 경우에는 기준액·고시금액·실구입금액 중 가장 낮은 금액의 10%에 해당하는 금액과 실구입금액과 기준액의 차액(실구입 \r\n" +
            "       금액이 고시금액보다 많은 경우에는 고시금액과 기준액의 차액을 말합니다)을 합산한 금액을 적습니다. 다만, 자세보조 용구는 실구입 금액을 적 \r\n" +
            "       습니다. \r\n" +
            "⑨ 청구금액: ⑤, ⑥ 및 ⑦ 중 가장 낮은 금액의 90%에 해당하는 금액을 적습니다. 다만, 국민건강보험법 시행령」 별표 2 제3호라목1)·2)에 해당하는 \r\n" +
            "       경감 대상자는 가장 낮은 금액의 100%에 해당하는 금액을 적습니다. \r\n" +
            "  <예시 1> 전동휠체어(기준액 2,090,000원, 고시금액 2,500,000원)를 2,000,000원에 구입한 경우 ⑧ 본인부담액란과 ⑨ 청구금액란에 기준액·고시 \r\n" +
            "       금액·실구입금액 중 가장 낮은 금액인 2,000,000원의 10%인 200,000원을 각각 적습니다. \r\n" +
            "  <예시 2> 전동스쿠터(기준액 1,670,000원, 고시금액 2,000,000원)를 2,000,000원에 구입한 경우 ⑧ 본인부담액란에는 기준액·고시금액·실구입금액 \r\n" +
            "       중 가장 낮은 금액인 1,670,000원의 10%인 167,000원과 실구입금액과 기준액의 차액인 330,000원을 합산한 금액인 497,000원을 적고, \r\n" +
            "       ⑨ 청구금액란에는 가장 낮은 금액인 1,670,000원의 90%인 1,503,000원을 적습니다. \r\n" +
            "  <예시 3> 체외용 인공후두(기준액 500,000원)를 550,000원에 구입한 경우 ⑧ 본인부담액란에는 기준액·고시금액·실구입금액 중 가장 낮은 금액인 \r\n" +
            "       500,000원의 10%인 50,000원과 실구입금액과 기준액의 차액인 50,000원을 합산한 금액인 100,000원을 적고, \r\n" +
            "       ⑨ 청구금액란에는 가장 낮은 금액인 500,000원의 90%인 450,000원을 적습니다. \r\n" +
            "⑩ 수령계좌: 보조기기 급여비를 받을 계좌를 선택하여 √ 표시를 하고, 금융기관명, 계좌번호, 예금주 성명, 주민(외국인)등록번호 또는 사업자등록 \r\n" +
            "     번호를 적습니다. \r\n" +
            "     ※ 예금주는 선택한 계좌에 따라 다음의 구분에 따른 사람이어야 합니다. \r\n" +
            "       - 가입자 또는 피부양자 계좌: 진료받은 사람, 진료받은 사람의 배우자 및 직계존비속, 진료받은 사람과 건강보험증을 같이 하거나 주민등록이 같\r\n" +
            "         이 되어 있는 형제자매 또는 직계비속의 배우자 \r\n" +
            "       - 보조기기 제조·대여·판매업소 계좌: 보조기기 제조·대여·판매업소의 법인 또는 대표자 \r\n" +
            "       - 진료받은 사람 본인의 요양비등 수급계좌(압류방지 계좌): 진료받은 사람 \r\n" +
            "     ※ 예금통장은 온라인 계좌입금이 가능한 예금통장이어야 합니다. \r\n" +
            "       (예시: 보통예금, 저축예금, 자유저축예금, 당좌예금 및 기업자유예금 등) \r\n" +
            "⑫ 청구인: 진료받은 사람, 진료받은 사람의 배우자 및 직계비속, 진료받은 사람과 건강보험증을 함께 하거나 주민등록이 함께 되어 있는 형제자매 \r\n" +
            "     또는 직계비속의 배우자여야 합니다. 이 경우 청구인은 본인의 이름을 적은 후 서명을 하거나 인장을 찍어야 하되, 청구인이 진료받은 사람으로서 \r\n" +
            "     제한능력자인 경우에는 법정대리인이 서명을 하거나 인장을 찍어 청구할 수 있습니다.");
        }
        sheet2.addMergedRegion(new CellRangeAddress(6, 6, 0, 33));

// 2 page end

// 3 page start

        // 워크시트 생성
        XSSFSheet sheet3 = workbook.createSheet("보청기 구매 표준계약서");
        sheet3.setMargin(HSSFSheet.TopMargin, 0.65);
        sheet3.setMargin(HSSFSheet.BottomMargin, 0.65);
        sheet3.setMargin(HSSFSheet.LeftMargin, 0.65);
        sheet3.setMargin(HSSFSheet.RightMargin, 0.65);
        // 행 생성
        XSSFRow row3 = null;
        // 셀 생성
        XSSFCell cell3;

        // 시트 정보 
        row3 = sheet3.createRow(0);
        row3.setHeight((short) 1000);
        cell3 = row3.createCell(0);
        for(int i=0; i<34; i++) {
            // 컬럼 폭
            sheet3.setColumnWidth(i, 620);
        }

        Font contractTitleFont = workbook.createFont();
        contractTitleFont.setFontHeight((short)320);    // 16pt
        contractTitleFont.setBold(true);
        contractTitleFont.setFontName("GungsuhChe");
        // 문서 제목 스타일 정의
        CellStyle contractTitleStyle = workbook.createCellStyle();
        contractTitleStyle.setAlignment(HorizontalAlignment.CENTER);
        contractTitleStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
        contractTitleStyle.setBorderTop(BorderStyle.NONE);
        contractTitleStyle.setBorderBottom(BorderStyle.NONE);
        contractTitleStyle.setBorderLeft(BorderStyle.NONE);
        contractTitleStyle.setBorderRight(BorderStyle.NONE);
        contractTitleStyle.setFont(contractTitleFont);

        Font contractTextFont = workbook.createFont();
        contractTextFont.setFontHeight((short)220);  // 11pt
        contractTextFont.setBold(false);
        contractTextFont.setFontName("BatangChe");

        // 문서 본문 스타일 정의
        CellStyle contractTextStyle = workbook.createCellStyle();
        contractTextStyle.setAlignment(HorizontalAlignment.LEFT);
        contractTextStyle.setVerticalAlignment(VerticalAlignment.TOP);
        contractTextStyle.setBorderTop(BorderStyle.NONE);
        contractTextStyle.setBorderBottom(BorderStyle.NONE);
        contractTextStyle.setBorderLeft(BorderStyle.NONE);
        contractTextStyle.setBorderRight(BorderStyle.NONE);
        contractTextStyle.setFont(contractTextFont);
        contractTextStyle.setWrapText(true);

        Font contractTextFontRed = workbook.createFont();
        contractTextFontRed.setFontHeight((short)220);  // 11pt
        contractTextFontRed.setBold(false);
        contractTextFontRed.setFontName("BatangChe");
        contractTextFontRed.setColor(Font.COLOR_RED);

        // 문서 본문 스타일 정의
        CellStyle contractTextStyleRed = workbook.createCellStyle();
        contractTextStyleRed.setAlignment(HorizontalAlignment.LEFT);
        contractTextStyleRed.setVerticalAlignment(VerticalAlignment.TOP);
        contractTextStyleRed.setBorderTop(BorderStyle.NONE);
        contractTextStyleRed.setBorderBottom(BorderStyle.NONE);
        contractTextStyleRed.setBorderLeft(BorderStyle.NONE);
        contractTextStyleRed.setBorderRight(BorderStyle.NONE);
        contractTextStyleRed.setFont(contractTextFontRed);
        contractTextStyleRed.setWrapText(true);

        Font contractTableFont = workbook.createFont();
        contractTableFont.setFontHeight((short)220);  // 11pt
        contractTableFont.setBold(true);
        contractTableFont.setFontName("BatangChe");

        // 표 좌측+상단 두꺼운 선 스타일 정의
        CellStyle conTableTopLeftStyle = workbook.createCellStyle();
        conTableTopLeftStyle.setAlignment(HorizontalAlignment.CENTER);
        conTableTopLeftStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        conTableTopLeftStyle.setBorderTop(BorderStyle.HAIR);
        conTableTopLeftStyle.setBorderBottom(BorderStyle.HAIR);
        conTableTopLeftStyle.setBorderLeft(BorderStyle.HAIR);
        conTableTopLeftStyle.setBorderRight(BorderStyle.HAIR);
        conTableTopLeftStyle.setFont(contractTableFont);
        conTableTopLeftStyle.setWrapText(true);

        // 제목 영역
        cell3 = row3.createCell(0);
        cell3.setCellStyle(contractTitleStyle);
        cell3.setCellValue("보청기 구매 표준계약서");
        sheet3.addMergedRegion(new CellRangeAddress(0, 0, 0, 33));

        // 공백 행
        row3 = sheet3.createRow(1);
        //row3.setHeight((short) 240);
        //row3.setHeight((short) 795); //1.40mm
        row3.setHeight((short) 680); // 1.20mm
        //row3.setHeight((short) 330); // 0.58mm
        sheet3.addMergedRegion(new CellRangeAddress(1, 1, 0, 33));

        // 첫문단
        row3 = sheet3.createRow(2);
        row3.setHeight((short) 900); // 0.58mm
        cell3 = row3.createCell(0);
        cell3.setCellStyle(contractTextStyle);
        cell3.setCellValue("  보청기에 대하여 보험급여를 신청하려는 가입자 또는 피부양자(이하 \"갑\"이라 함)와 국민건강보험공단(이하 \"공단\"이라 함)에 보청기판매업소로 등록한 자(이하 \"을\"이라 함)는 다음과 같이 보청기 구매 계약을 체결한다.");
        sheet3.addMergedRegion(new CellRangeAddress(2, 2, 0, 33));

        // 공백 행
        row3 = sheet3.createRow(3);
        //row3.setHeight((short) 240);
        row3.setHeight((short) 400); // 0.58mm
        //row3.setHeight((short) 795); //1.40mm
        //row3.setHeight((short) 680); // 1.20mm
        sheet3.addMergedRegion(new CellRangeAddress(3, 3, 0, 33));

        // 제1조
        row3 = sheet3.createRow(4);
        row3.setHeight((short) 400); // 0.58mm
        cell3 = row3.createCell(0);
        cell3.setCellStyle(contractTextStyle);
        cell3.setCellValue("제1조(구매 대상) 이 계약에 따라 갑이 을에게서 구입하는 보청기는 다음과 같다.");
        sheet3.addMergedRegion(new CellRangeAddress(4, 4, 0, 33));

        // 제1조 표
        row3 = sheet3.createRow(5);
        row3.setHeight((short) 420); // 0.58mm
        for(int i=0; i<5; i++) {
            cell3 = row3.createCell(i);
            cell3.setCellStyle(conTableTopLeftStyle);
            cell3.setCellValue("구분");
        }
        sheet3.addMergedRegion(new CellRangeAddress(5, 6, 0, 4));
        for(int i=5; i<14; i++) {
            cell3 = row3.createCell(i);
            cell3.setCellStyle(conTableTopLeftStyle);
            cell3.setCellValue("모델명");
        }
        sheet3.addMergedRegion(new CellRangeAddress(5, 6, 5, 13));
        for(int i=14; i<24; i++) {
            cell3 = row3.createCell(i);
            cell3.setCellStyle(conTableTopLeftStyle);
            cell3.setCellValue("구매금액");
        }
        sheet3.addMergedRegion(new CellRangeAddress(5, 5, 14, 23));
        for(int i=24; i<29; i++) {
            cell3 = row3.createCell(i);
            cell3.setCellStyle(conTableTopLeftStyle);
            cell3.setCellValue("제조사");
        }
        sheet3.addMergedRegion(new CellRangeAddress(5, 6, 24, 28));
        for(int i=29; i<34; i++) {
            cell3 = row3.createCell(i);
            cell3.setCellStyle(conTableTopLeftStyle);
            cell3.setCellValue("형태");
        }
        sheet3.addMergedRegion(new CellRangeAddress(5, 6, 29, 33));

        row3 = sheet3.createRow(6);
        row3.setHeight((short) 420); // 0.58mm
        for(int i=0; i<5; i++) {
            cell3 = row3.createCell(i);
            cell3.setCellStyle(conTableTopLeftStyle);
            cell3.setCellValue("구분");
        }
        //sheet3.addMergedRegion(new CellRangeAddress(5, 6, 0, 4));
        for(int i=14; i<19; i++) {
            cell3 = row3.createCell(i);
            cell3.setCellStyle(conTableTopLeftStyle);
            cell3.setCellValue("단가");
        }
        sheet3.addMergedRegion(new CellRangeAddress(6, 6, 14, 18));
        for(int i=19; i<24; i++) {
            cell3 = row3.createCell(i);
            cell3.setCellStyle(conTableTopLeftStyle);
            cell3.setCellValue("계");
        }
        sheet3.addMergedRegion(new CellRangeAddress(6, 6, 19, 23));
        for(int i=24; i<29; i++) {
            cell3 = row3.createCell(i);
            cell3.setCellStyle(conTableTopLeftStyle);
            cell3.setCellValue("제조사");
        }
        for(int i=29; i<34; i++) {
            cell3 = row3.createCell(i);
            cell3.setCellStyle(conTableTopLeftStyle);
            cell3.setCellValue("형태");
        }

        row3 = sheet3.createRow(7);
        row3.setHeight((short) 600); // 0.58mm
        for(int i=0; i<5; i++) {
            cell3 = row3.createCell(i);
            cell3.setCellStyle(conTableTopLeftStyle);
            cell3.setCellValue("오른쪽 귀");
        }
        sheet3.addMergedRegion(new CellRangeAddress(7, 7, 0, 4));
        for(int i=5; i<14; i++) {
            cell3 = row3.createCell(i);
            cell3.setCellStyle(conTableTopLeftStyle);
            cell3.setCellValue("");
        }
        sheet3.addMergedRegion(new CellRangeAddress(7, 7, 5, 13));
        for(int i=14; i<19; i++) {
            cell3 = row3.createCell(i);
            cell3.setCellStyle(conTableTopLeftStyle);
            cell3.setCellValue("");
        }
        sheet3.addMergedRegion(new CellRangeAddress(7, 7, 14, 18));
        for(int i=19; i<24; i++) {
            cell3 = row3.createCell(i);
            cell3.setCellStyle(conTableTopLeftStyle);
            cell3.setCellValue("");
        }
        sheet3.addMergedRegion(new CellRangeAddress(7, 7, 19, 23));
        for(int i=24; i<29; i++) {
            cell3 = row3.createCell(i);
            cell3.setCellStyle(conTableTopLeftStyle);
            cell3.setCellValue("");
        }
        sheet3.addMergedRegion(new CellRangeAddress(7, 7, 24, 28));
        for(int i=29; i<34; i++) {
            cell3 = row3.createCell(i);
            cell3.setCellStyle(conTableTopLeftStyle);
            cell3.setCellValue("");
        }
        sheet3.addMergedRegion(new CellRangeAddress(7, 7, 29, 33));

        row3 = sheet3.createRow(8);
        row3.setHeight((short) 600); // 0.58mm
        for(int i=0; i<5; i++) {
            cell3 = row3.createCell(i);
            cell3.setCellStyle(conTableTopLeftStyle);
            cell3.setCellValue("왼쪽 귀");
        }
        sheet3.addMergedRegion(new CellRangeAddress(8, 8, 0, 4));
        for(int i=5; i<14; i++) {
            cell3 = row3.createCell(i);
            cell3.setCellStyle(conTableTopLeftStyle);
            cell3.setCellValue("");
        }
        sheet3.addMergedRegion(new CellRangeAddress(8, 8, 5, 13));
        for(int i=14; i<19; i++) {
            cell3 = row3.createCell(i);
            cell3.setCellStyle(conTableTopLeftStyle);
            cell3.setCellValue("");
        }
        sheet3.addMergedRegion(new CellRangeAddress(8, 8, 14, 18));
        for(int i=19; i<24; i++) {
            cell3 = row3.createCell(i);
            cell3.setCellStyle(conTableTopLeftStyle);
            cell3.setCellValue("");
        }
        sheet3.addMergedRegion(new CellRangeAddress(8, 8, 19, 23));
        for(int i=24; i<29; i++) {
            cell3 = row3.createCell(i);
            cell3.setCellStyle(conTableTopLeftStyle);
            cell3.setCellValue("");
        }
        sheet3.addMergedRegion(new CellRangeAddress(8, 8, 24, 28));
        for(int i=29; i<34; i++) {
            cell3 = row3.createCell(i);
            cell3.setCellStyle(conTableTopLeftStyle);
            cell3.setCellValue("");
        }
        sheet3.addMergedRegion(new CellRangeAddress(8, 8, 29, 33));

        // 공백 행
        row3 = sheet3.createRow(9);
        //row3.setHeight((short) 240);
        row3.setHeight((short) 400); // 0.58mm
        //row3.setHeight((short) 795); //1.40mm
        //row3.setHeight((short) 680); // 1.20mm
        sheet3.addMergedRegion(new CellRangeAddress(9, 9, 0, 33));

        // 제2조
        row3 = sheet3.createRow(10);
        row3.setHeight((short) 2500); // 0.58mm

        cell3 = row3.createCell(0);
        cell3.setCellStyle(contractTextStyle);
        cell3.setCellValue("제2조(적합관리 서비스) 을은 보청기 성능의 유지·관리를 위하여 보청기를 구매한 날로부터 \n" +
        "  갑에게 다음의 각 호의 서비스를 한다. \n" +
        " 1. 보청기 사용 및 관리에 관한 상담 \n" +
        " 2. 청각평가 \n" +
        " 3. 보청기 음량조절, 귀꽂이의 변형·정비 \n" +
        " 4. 보청기, 부속품, 보조장치의 변형 및 정비 \n" +
        " 5. 청능훈련 \n" +
        " 6. 상호작용과 대화를 통한 서비스 제공 \n" +
        " 7. 그 밖에 보청기 성능의 유지·관리를 위해 필요한 사항으로서 갑과 을이 따로 정하는 사항");
        sheet3.addMergedRegion(new CellRangeAddress(10, 10, 0, 33));

        // 공백 행
        row3 = sheet3.createRow(11);
        //row3.setHeight((short) 240);
        row3.setHeight((short) 400); // 0.58mm
        //row3.setHeight((short) 795); //1.40mm
        //row3.setHeight((short) 680); // 1.20mm
        sheet3.addMergedRegion(new CellRangeAddress(11, 11, 0, 33));

        // 제3조
        row3 = sheet3.createRow(12);
        row3.setHeight((short) 600); // 0.58mm
        cell3 = row3.createCell(0);
        cell3.setCellStyle(contractTextStyle);
        cell3.setCellValue("제3조(갑의 의무) 갑은 보청기 구입비용 및 제2조에 따른 적합관리 비용을 을이 공단에게서 \n" +
        " 지급받도록 하는 경우에는 그 지급이 원활하게 이루어질 수 있도록 적극 협조해야 한다.");
        sheet3.addMergedRegion(new CellRangeAddress(12, 12, 0, 33));

        // 공백 행
        row3 = sheet3.createRow(13);
        //row3.setHeight((short) 240);
        row3.setHeight((short) 400); // 0.58mm
        //row3.setHeight((short) 795); //1.40mm
        //row3.setHeight((short) 680); // 1.20mm
        sheet3.addMergedRegion(new CellRangeAddress(13, 13, 0, 33));

        // 제4조
        row3 = sheet3.createRow(14);
        row3.setHeight((short) 2400); // 0.58mm
        cell3 = row3.createCell(0);
        cell3.setCellStyle(contractTextStyle);
        cell3.setCellValue("제4조(을의 의무) ① 을은 공단에 등록된 보청기 판매업소로서 관련 규정을 준수해야 한다. \n" +
        " ② 을은 보청기와 보청기를 관리하는 시설·장비에 대하여 적절한 위생관리를 실시해야 한다. \n" +
        " ③ 을은 다음 각 호의 어느 하나의 사유가 발생하여 제2조에 따른 서비스 및 이 조에 다른 의 \n" +
        "    무를 이행할 수 없는 경우에는 갑에게 그 사실을 알려 갑이 적합관리를 받는데 불편함이 \n" +
        "    없도록 하여야 한다. \n" +
        "  1. 휴업 또는 폐업 \n" +
        "  2. 사업장의 이전 \n" +
        "  3. 운영인력, 장비 등의 결여 \n" +
        "  4. 그 밖에 해당 서비스 및 의무를 이행할 수 없는 부득이한 사유");
        sheet3.addMergedRegion(new CellRangeAddress(14, 14, 0, 33));

        // 공백 행
        row3 = sheet3.createRow(15);
        //row3.setHeight((short) 240);
        row3.setHeight((short) 400); // 0.58mm
        //row3.setHeight((short) 795); //1.40mm
        //row3.setHeight((short) 680); // 1.20mm
        sheet3.addMergedRegion(new CellRangeAddress(15, 15, 0, 33));

        // 제5조 -1
        row3 = sheet3.createRow(16);
        row3.setHeight((short) 300); // 0.58mm
        cell3 = row3.createCell(0);
        cell3.setCellStyle(contractTextStyle);
        cell3.setCellValue("제5조(계약의 해제) ① 갑은 다음 각 호의 어느 하나에 해당하는 경우 계약을 해제할 수 있다.");
        sheet3.addMergedRegion(new CellRangeAddress(16, 16, 0, 33));
        // 제5조 -2
        row3 = sheet3.createRow(17);
        row3.setHeight((short) 300); // 0.58mm
        cell3 = row3.createCell(0);
        cell3.setCellStyle(contractTextStyleRed);
        cell3.setCellValue("  1. 계약체결 시 보청기에 대해 을이 정보를 정확하게 제공되지 않은 경우");
        sheet3.addMergedRegion(new CellRangeAddress(17, 17, 0, 33));
        // 제5조 -3
        row3 = sheet3.createRow(18);
        row3.setHeight((short) 300); // 0.58mm
        cell3 = row3.createCell(0);
        cell3.setCellStyle(contractTextStyle);
        cell3.setCellValue("  2. 을의 호객행위로 계약을 체결한 경우");
        sheet3.addMergedRegion(new CellRangeAddress(18, 18, 0, 33));
        // 제5조 -4
        row3 = sheet3.createRow(19);
        row3.setHeight((short) 300); // 0.58mm
        cell3 = row3.createCell(0);
        cell3.setCellStyle(contractTextStyleRed);
        cell3.setCellValue("  3. 보청기 착용 후 청력개선 효과가 없어 갑이 검수확인을 받지 못한 경우");
        sheet3.addMergedRegion(new CellRangeAddress(19, 19, 0, 33));

// 3 page end

// 4 page start

        // 워크시트 생성
        XSSFSheet sheet4 = workbook.createSheet("보청기 구매 표준계약서2");
        sheet4.setMargin(HSSFSheet.TopMargin, 0.65);
        sheet4.setMargin(HSSFSheet.BottomMargin, 0.65);
        sheet4.setMargin(HSSFSheet.LeftMargin, 0.65);
        sheet4.setMargin(HSSFSheet.RightMargin, 0.65);
        // 행 생성
        XSSFRow row4 = null;
        // 셀 생성
        XSSFCell cell4;

        // 시트 정보 
        row4 = sheet4.createRow(0);
        row4.setHeight((short) 1000);
        cell4 = row4.createCell(0);
        for(int i=0; i<34; i++) {
            // 컬럼 폭
            sheet4.setColumnWidth(i, 620);
        }

        // 공백 행
        row4 = sheet4.createRow(0);
        //row4.setHeight((short) 240);
        row4.setHeight((short) 400); // 0.58mm
        //row4.setHeight((short) 795); //1.40mm
        //row4.setHeight((short) 680); // 1.20mm
        sheet4.addMergedRegion(new CellRangeAddress(0, 0, 0, 33));

        // 제5조 -5
        row4 = sheet4.createRow(1);
        row4.setHeight((short) 1200); // 0.58mm
        cell4 = row4.createCell(0);
        cell4.setCellStyle(contractTextStyle);
        cell4.setCellValue(" ② 제1항에 따라 계약이 해제된 경우 갑은 즉시 보청기를 반환하고 을은 총 구매금액을 반환\n" +
        " 해야 한다. 다만, 을이 대금 중 일부를 공단에게서 지급받은 경우에는 갑이 지불한 금액은 갑 \n" +
        " 에게, 공단이 지급한 금액은 공단에 각각 반환해야 하며, 을은 반환 후 그 사실을 갑에게 통 \n" +
        " 지해야 한다.");
        sheet4.addMergedRegion(new CellRangeAddress(1, 1, 0, 33));

        // 공백 행
        row4 = sheet4.createRow(2);
        //row4.setHeight((short) 240);
        row4.setHeight((short) 400); // 0.58mm
        //row4.setHeight((short) 795); //1.40mm
        //row4.setHeight((short) 680); // 1.20mm
        sheet4.addMergedRegion(new CellRangeAddress(2, 2, 0, 33));

        // 제6조 -1
        row4 = sheet4.createRow(3);
        row4.setHeight((short) 600); // 0.58mm
        cell4 = row4.createCell(0);
        cell4.setCellStyle(contractTextStyle);
        cell4.setCellValue("제6조(계약의 해지) ① 갑은 다음 각 호에 어느 하나의 사유가 있는 경우에는 계약을 해지할 \n" +
        "   수 있다.");
        sheet4.addMergedRegion(new CellRangeAddress(3, 3, 0, 33));

        // 제6조 -2
        row4 = sheet4.createRow(4);
        row4.setHeight((short) 300); // 0.58mm
        cell4 = row4.createCell(0);
        cell4.setCellStyle(contractTextStyleRed);
        cell4.setCellValue("  1. 을이 사후 적합관리 증명서를 발급해 주지 않는 경우");
        sheet4.addMergedRegion(new CellRangeAddress(4, 4, 0, 33));

        // 제6조 -3
        row4 = sheet4.createRow(5);
        row4.setHeight((short) 300); // 0.58mm
        cell4 = row4.createCell(0);
        cell4.setCellStyle(contractTextStyleRed);
        cell4.setCellValue("  2. 을이 「개인정보 보호법」을 위반하여 갑의 개인정보를 처리한 경우");
        sheet4.addMergedRegion(new CellRangeAddress(5, 5, 0, 33));

        // 제6조 -4
        row4 = sheet4.createRow(6);
        row4.setHeight((short) 300); // 0.58mm
        cell4 = row4.createCell(0);
        cell4.setCellStyle(contractTextStyleRed);
        cell4.setCellValue("  3. 갑의 이사 등 부득이한 사유로 적합관리를 받을 수 없게 된 경우");
        sheet4.addMergedRegion(new CellRangeAddress(6, 6, 0, 33));

        // 제6조 -5
        row4 = sheet4.createRow(7);
        row4.setHeight((short) 300); // 0.58mm
        cell4 = row4.createCell(0);
        cell4.setCellStyle(contractTextStyle);
        cell4.setCellValue("  ② 갑은 제1항에 따라 계약을 해지하려면 을에게 그 의사를 통지해야 한다.");
        sheet4.addMergedRegion(new CellRangeAddress(7, 7, 0, 33));

        // 제6조 -6
        row4 = sheet4.createRow(8);
        row4.setHeight((short) 600); // 0.58mm
        cell4 = row4.createCell(0);
        cell4.setCellStyle(contractTextStyle);
        cell4.setCellValue("  ③ 갑은 을의 고의 또는 중대한 과실로 손해를 입은 경우에는 제2항에도 불구하고 사전 통지 \n" +
        "  없이 일방적인 의사표시로 계약을 해지할 수 있다.");
        sheet4.addMergedRegion(new CellRangeAddress(8, 8, 0, 33));

        // 공백 행
        row4 = sheet4.createRow(9);
        //row4.setHeight((short) 240);
        row4.setHeight((short) 400); // 0.58mm
        //row4.setHeight((short) 795); //1.40mm
        //row4.setHeight((short) 680); // 1.20mm
        sheet4.addMergedRegion(new CellRangeAddress(9, 9, 0, 33));

        // 제7조
        row4 = sheet4.createRow(10);
        row4.setHeight((short) 1500); // 0.58mm
        cell4 = row4.createCell(0);
        cell4.setCellStyle(contractTextStyle);
        cell4.setCellValue("제7조(개인정보 보호) ① 을은 보청기 판매와 적합관리 과정에서 알게 된 갑의 개인정보를 관 \n" +
        "  계 규정에 따라 보호해야 한다. \n" +
        "  ② 을은 제2조에 따른 서비스 제공기간이 종료되면 갑에게 지체 없이 적합관리 기간 동안의 적합관리기록의 원본을 제공하고 \n" +
        "  을 본인이 해당 기록을 보관해서는 아니 된다.");
        sheet4.addMergedRegion(new CellRangeAddress(10, 10, 0, 33));

        // 공백 행
        row4 = sheet4.createRow(11);
        //row4.setHeight((short) 240);
        row4.setHeight((short) 400); // 0.58mm
        //row4.setHeight((short) 795); //1.40mm
        //row4.setHeight((short) 680); // 1.20mm
        sheet4.addMergedRegion(new CellRangeAddress(11, 11, 0, 33));

        // 제8조
        row4 = sheet4.createRow(12);
        row4.setHeight((short) 2400); // 0.58mm
        cell4 = row4.createCell(0);
        cell4.setCellStyle(contractTextStyle);
        cell4.setCellValue("제8조(보칙) ① 이 계약서에서 정하지 않은 사항에 대해서는 「국민건강보호법」, 「소비자기\n" +
        "  본법」,「약관의 규제에 관한 법률」, 「할부거래에 관한 법률」,「방문판매 등에 관한 법 \n" +
        "  률」,「전자상거래 등에서의 소비자보호에 관한 법률」,「민법」 등 관계 법령에 따르며, \n" +
        "  \"갑\"과 \"을\"이 개별적으로 약정한 사항이 있는 경우에는 해당 관계 법령 내 강행규정에 반하 \n" +
        "  지 않는 한 그 약정한 바에 따른다. \n" +
        "  ② 위 계약 체결을 증명하고 제반 의무를 성실히 수행하기 위하여 본 계약서를 2부 작성하여 \n" +
        "  서명 날인 후 갑과 을이 각각 1부씩 보관한다.");
        sheet4.addMergedRegion(new CellRangeAddress(12, 12, 0, 33));


        // 문서 본문 스타일 정의
        CellStyle contractTextRightStyle = workbook.createCellStyle();
        contractTextRightStyle.setAlignment(HorizontalAlignment.RIGHT);
        contractTextRightStyle.setVerticalAlignment(VerticalAlignment.TOP);
        contractTextRightStyle.setBorderTop(BorderStyle.NONE);
        contractTextRightStyle.setBorderBottom(BorderStyle.NONE);
        contractTextRightStyle.setBorderLeft(BorderStyle.NONE);
        contractTextRightStyle.setBorderRight(BorderStyle.NONE);
        contractTextRightStyle.setFont(contractTextFont);
        contractTextRightStyle.setWrapText(true);

        // 년월일
        row4 = sheet4.createRow(13);
        row4.setHeight((short) 600); // 0.58mm
        cell4 = row4.createCell(0);
        cell4.setCellStyle(contractTextRightStyle);
        cell4.setCellValue("         년       월       일");
        sheet4.addMergedRegion(new CellRangeAddress(13, 13, 0, 33));

// 4 page end


        // 입력된 내용 파일로 쓰기
        //File file = new File("C:\\excel\\testWrite.xlsx");
        String userName = System.getProperty("user.home");
        File file = new File(userName+"\\Downloads\\testWrite.xlsx");
        FileOutputStream fos = null;

        try {
            fos = new FileOutputStream(file);
            workbook.write(fos);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if(workbook!=null) workbook.close();
                if(fos!=null) fos.close();

            } catch (IOException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
        }

        //return "open /user/download/";
        return "/getForm1";
    }

    // 헥스코드(#009900)를 바이트로 변환.
    public XSSFColor StringToHexColor(String rgbS) throws Exception {
        //String rgbS = "D9D9D9";
        byte[] rgbB = Hex.decodeHex(rgbS);
        XSSFColor color = new XSSFColor(rgbB, null);
        return color;
    }

}
