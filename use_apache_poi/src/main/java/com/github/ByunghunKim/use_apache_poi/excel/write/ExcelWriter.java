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
