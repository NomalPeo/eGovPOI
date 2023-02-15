package egovframework.board.controller;

import java.io.IOException;

import javax.servlet.http.HttpServletResponse;import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

@Controller
public class boardController {

	@RequestMapping(value="/main.do")
	public String main() {
		return "main";
	}

	@RequestMapping("/excel.do")
    public void excelDownload(HttpServletResponse response) throws IOException {
//        Workbook wb = new HSSFWorkbook();
		
        Workbook wb = new XSSFWorkbook();
        CellStyle style = wb.createCellStyle();
        Font font = wb.createFont();
        font.setFontHeightInPoints((short)12);
        style.setFont(font);
        Sheet sheet = wb.createSheet("첫번째 시트");
        Row row = null;
        Cell cell = null;
        int rowNum = 4;
        // Header
        row = sheet.createRow(rowNum++); //4
        sheet.addMergedRegion( new CellRangeAddress(4,4,0,7));
        for(int i=0; i<8; i++) {
        	CellStyle stylebalck1 = wb.createCellStyle();
        	cell = row.createCell(i);
        	if(i==0) {
        		cell.setCellValue("(08507) 서울시 금천구 가산디지털1로 168 우림라이온스밸리 A동 8층 ㈜아사달 (전화: 1544-8442, 팩스: 02-2026-2008)");
        		stylebalck1.setAlignment((short)2);
        	}
        	stylebalck1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setBorderTop(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setBorderRight(HSSFCellStyle.BORDER_THIN);
            cell.setCellStyle(stylebalck1);
        }
        
        row = sheet.createRow(rowNum++); // 5
        cell = row.createCell(0);
    	CellStyle stylebalck2 = wb.createCellStyle();
    	Font font1 = wb.createFont();
    	font1.setFontHeightInPoints((short)28);
    	cell.setCellValue("가격 명세서");
    	stylebalck2.setFont(font1);
        sheet.addMergedRegion( new CellRangeAddress(5,5,0,7));
        style.setAlignment((short)2);
        cell.setCellStyle(style);
        
        row = sheet.createRow(rowNum++);
        sheet.addMergedRegion( new CellRangeAddress(6,6,4,7));
        for(int i=4; i<8; i++) {
        	cell = row.createCell(i);
            CellStyle stylebalck = wb.createCellStyle();
        	if(i==4) {
                cell.setCellValue("작성일 : 2021년 12월 21일");
                stylebalck.setAlignment((short)2);
        	}
            stylebalck.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            stylebalck.setBorderTop(HSSFCellStyle.BORDER_THIN);
            stylebalck.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            stylebalck.setBorderRight(HSSFCellStyle.BORDER_THIN);
            cell.setCellStyle(stylebalck);
        	
        }
        
        row = sheet.createRow(rowNum++);
        for(int i=0; i<8; i++) {
            CellStyle stylebalck1 = wb.createCellStyle();
            cell = row.createCell(i);
        	String[] name = new String[] {"수신자","발주기관","서울시","공급자","회사명","㈜아사달시스템"};
        	if(i==7 || i ==3) {
        	}else if(i>2) {
        		cell.setCellValue(name[i-1]);
        	}else {
        		cell.setCellValue(name[i]);
        	}
            sheet.addMergedRegion( new CellRangeAddress(7,9,0,0));
            sheet.addMergedRegion( new CellRangeAddress(7,7,2,3));
            sheet.addMergedRegion( new CellRangeAddress(7,9,4,4));
            sheet.addMergedRegion( new CellRangeAddress(7,7,6,7));
            stylebalck1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setBorderTop(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setBorderRight(HSSFCellStyle.BORDER_THIN);
            cell.setCellStyle(stylebalck1);
        }
        
        row = sheet.createRow(rowNum++);
        for(int i=0; i<8; i++) {
            CellStyle stylebalck1 = wb.createCellStyle();
            cell = row.createCell(i);
        	String[] name = new String[] {null,"대표자","오세훈",null,null,"대표자","이종원(인)",null};
        	if(i==3||i==7||i==0||i==4) {
        		
        	}else{
        		cell.setCellValue(name[i]);
        	}
            sheet.addMergedRegion( new CellRangeAddress(8,8,2,3));
            sheet.addMergedRegion( new CellRangeAddress(8,8,6,7));
            stylebalck1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setBorderTop(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setBorderRight(HSSFCellStyle.BORDER_THIN);
            cell.setCellStyle(stylebalck1);
        }
        row = sheet.createRow(rowNum++);
        for(int i=0; i<8; i++) {
            CellStyle stylebalck1 = wb.createCellStyle();
            cell = row.createCell(i);
        	String[] name = new String[] {null,"담당자","오세훈 시장",null,null,"담당자","진가연 사원",null};
        	if(i==3||i==7||i==0||i==4) {
        		
        	}else{
        		cell.setCellValue(name[i]);
        	}
            sheet.addMergedRegion( new CellRangeAddress(9,9,2,3));
            sheet.addMergedRegion( new CellRangeAddress(9,9,6,7));
            stylebalck1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setBorderTop(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setBorderRight(HSSFCellStyle.BORDER_THIN);
            cell.setCellStyle(stylebalck1);
        }
        
        row = sheet.createRow(rowNum++);
        row = sheet.createRow(rowNum++);
        
        for(int i=0; i<8; i++) {
        	cell = row.createCell(i);
        	if(i==0) {
                cell.setCellValue("한국저작권 위원회");
                sheet.addMergedRegion( new CellRangeAddress(11,11,0,7));
        	}
        	CellStyle stylebalck1 = wb.createCellStyle();
            stylebalck1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setBorderTop(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setBorderRight(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setAlignment((short)2);
            cell.setCellStyle(stylebalck1);
        }
        
        row = sheet.createRow(rowNum++);
        row = sheet.createRow(rowNum++);
        row = sheet.createRow(rowNum++);
        for(int i=0; i<8; i++) {
            cell = row.createCell(i);
        	if(i==0) {cell.setCellValue("구분");
        	}else if(i==1){cell.setCellValue("산출내역");
        	}else if(i==6) {cell.setCellValue("금액(원)");
        	}else if(i==7) {cell.setCellValue("메모");}
            sheet.addMergedRegion( new CellRangeAddress(14,15,0,0));
            sheet.addMergedRegion( new CellRangeAddress(14,14,1,5));
            sheet.addMergedRegion( new CellRangeAddress(14,15,6,6));
            sheet.addMergedRegion( new CellRangeAddress(14,15,7,7));
            CellStyle stylebalck1 = wb.createCellStyle();
            stylebalck1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setBorderTop(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setBorderRight(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setAlignment((short)2);
            cell.setCellStyle(stylebalck1);
        }
        
        row = sheet.createRow(rowNum++);
        for(int i=0; i<8; i++) {
            cell = row.createCell(i);
            String [] name = new String[] {"역할","기술등급","성명","단가(원)","투입인력"};
            if(0<i && i<6) {
            	cell.setCellValue(name[i-1]);
            }
            CellStyle stylebalck1 = wb.createCellStyle();
            stylebalck1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setBorderTop(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setBorderRight(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setAlignment((short)2);
            cell.setCellStyle(stylebalck1);
        }
        
        row = sheet.createRow(rowNum++);
        for(int i=1; i<8; i++) {
            cell = row.createCell(i);
            if(i==0) {
                cell.setCellValue("내부 인건비");
            }
            sheet.addMergedRegion( new CellRangeAddress(16,16,0,7));
            CellStyle stylebalck1 = wb.createCellStyle();
            stylebalck1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setBorderTop(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setBorderRight(HSSFCellStyle.BORDER_THIN);
            stylebalck1.setAlignment((short)2);
            cell.setCellStyle(stylebalck1);
        }
        
        
        
        
        row = sheet.createRow(rowNum++);
        cell = row.createCell(0);
        sheet.addMergedRegion( new CellRangeAddress(17,28,0,0));
        cell = row.createCell(1);
        cell.setCellValue("개발/구축");
        sheet.addMergedRegion( new CellRangeAddress(17,17,1,7));

        row = sheet.createRow(rowNum++);
        cell = row.createCell(1);
        cell.setCellValue("총괄책임");
        cell = row.createCell(2);
        cell.setCellValue("특급기술자");
        cell = row.createCell(3);
        cell.setCellValue("총괄자");
        cell = row.createCell(4);
        cell.setCellValue(7000000);
        cell = row.createCell(5);
        cell.setCellValue(0.5);
        cell = row.createCell(6);
        cell.setCellValue("=E19*F19");
        
        row = sheet.createRow(rowNum++);
        cell = row.createCell(1);
        cell.setCellValue("PM");
        cell = row.createCell(2);
        cell.setCellValue("특급기술자");
        cell = row.createCell(3);
        cell.setCellValue("PM");
        cell = row.createCell(4);
        cell.setCellValue(7000000);
        cell = row.createCell(5);
        cell.setCellValue(4);
        cell = row.createCell(6);
        cell.setCellValue("=E20*F20");
        
        row = sheet.createRow(rowNum++);
        cell = row.createCell(1);
        cell.setCellValue("개발자");
        cell = row.createCell(2);
        cell.setCellValue("중급기술자");
        cell = row.createCell(3);
        cell.setCellValue("개발자1");
        cell = row.createCell(4);
        cell.setCellValue(6500000);
        cell = row.createCell(5);
        cell.setCellValue(1);
        cell = row.createCell(6);
        cell.setCellValue("=E21*F21");
        
        row = sheet.createRow(rowNum++);
        cell = row.createCell(1);
        cell.setCellValue("SE");
        cell = row.createCell(2);
        cell.setCellValue("고급기술자");
        cell = row.createCell(3);
        cell.setCellValue("개발자2");
        cell = row.createCell(4);
        cell.setCellValue(5000000);
        cell = row.createCell(5);
        cell.setCellValue(1);
        cell = row.createCell(6);
        cell.setCellValue("=E22*F22");
        
        row = sheet.createRow(rowNum++);
        cell = row.createCell(1);
        cell.setCellValue("기획자");
        cell = row.createCell(2);
        cell.setCellValue("중급기술자");
        cell = row.createCell(3);
        cell.setCellValue("웹디자이너1");
        cell = row.createCell(4);
        cell.setCellValue(5000000);
        cell = row.createCell(5);
        cell.setCellValue(1);
        cell = row.createCell(6);
        cell.setCellValue("=E23*F23");
        
        row = sheet.createRow(rowNum++);
        cell = row.createCell(1);
        cell.setCellValue("웹디자이너");
        cell = row.createCell(2);
        cell.setCellValue("중급기술자");
        cell = row.createCell(3);
        cell.setCellValue("웹디자이너1");
        cell = row.createCell(4);
        cell.setCellValue(6000000);
        cell = row.createCell(5);
        cell.setCellValue(1);
        cell = row.createCell(6);
        cell.setCellValue("=E24*F24");
        
        row = sheet.createRow(rowNum++);
        cell = row.createCell(1);
        cell.setCellValue("퍼블리셔");
        cell = row.createCell(2);
        cell.setCellValue("초급기술자");
        cell = row.createCell(3);
        cell.setCellValue("퍼블리셔1");
        cell = row.createCell(4);
        cell.setCellValue(5000000);
        cell = row.createCell(5);
        cell.setCellValue(1);
        cell = row.createCell(6);
        cell.setCellValue("=E25*F25");

        row = sheet.createRow(rowNum++);
        cell = row.createCell(1);
        cell.setCellValue("품질관리");
        cell = row.createCell(2);
        cell.setCellValue("고급기술자");
        cell = row.createCell(3);
        cell.setCellValue("품질관리자1");
        cell = row.createCell(4);
        cell.setCellValue(7000000);
        cell = row.createCell(5);
        cell.setCellValue(1);
        cell = row.createCell(6);
        cell.setCellValue("=E26*F26");
        
        row = sheet.createRow(rowNum++);
        cell = row.createCell(1);
        cell.setCellValue("소계");
        sheet.addMergedRegion( new CellRangeAddress(26,26,1,4));
        cell = row.createCell(5);
        cell.setCellValue("10.5");
        cell = row.createCell(6);
        cell.setCellValue("=SUM(G19:G26)");
        
        row = sheet.createRow(rowNum++);
        cell = row.createCell(1);
        cell.setCellValue("재경비");
        sheet.addMergedRegion( new CellRangeAddress(27,27,1,5));
        cell = row.createCell(6);
        cell.setCellValue("=INT(G27*(B28/100))");
        
        row = sheet.createRow(rowNum++);
        cell = row.createCell(1);
        cell.setCellValue("내부 인건비 소계");
        sheet.addMergedRegion( new CellRangeAddress(28,28,1,5));
        cell = row.createCell(6);
        cell.setCellValue("=G27+G28");
        
        row = sheet.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("외주비");
        sheet.addMergedRegion( new CellRangeAddress(29,29,0,7));

        row = sheet.createRow(rowNum++);
        cell = row.createCell(0);
        sheet.addMergedRegion( new CellRangeAddress(30,40,0,0));
        cell = row.createCell(1);
        cell.setCellValue("제품명");
        sheet.addMergedRegion( new CellRangeAddress(30,30,1,2));
        cell = row.createCell(3);
        cell.setCellValue("업체명");
        cell = row.createCell(4);
        cell.setCellValue("단가(원)");
        cell = row.createCell(5);
        cell.setCellValue("수량(식)");
        cell = row.createCell(6);
        cell.setCellValue("금액");
        
        row = sheet.createRow(rowNum++);
        cell = row.createCell(1);
        cell.setCellValue("버스티켓");
        sheet.addMergedRegion( new CellRangeAddress(31,31,1,2));
        cell = row.createCell(3);
        cell.setCellValue("동부고속");
        cell = row.createCell(4);
        cell.setCellValue(1520000);
        cell = row.createCell(5);
        cell.setCellValue(1);
        cell = row.createCell(6);
        cell.setCellValue("=E32*F32");

        row = sheet.createRow(rowNum++);
        cell = row.createCell(1);
        sheet.addMergedRegion( new CellRangeAddress(32,32,1,2));
        cell = row.createCell(3);
        cell = row.createCell(4);
        cell = row.createCell(5);
        cell = row.createCell(6);
        cell = row.createCell(7);
        cell.setCellValue("위 합산");
        
        for(int i=33; i<39; i++) {
        	row = sheet.createRow(rowNum++);
            cell = row.createCell(1);
            sheet.addMergedRegion( new CellRangeAddress(i,i,1,2));
            cell = row.createCell(3);
            cell = row.createCell(4);
            cell = row.createCell(5);
            cell = row.createCell(6);
            cell = row.createCell(7);
        }


        row = sheet.createRow(rowNum++);
        row = sheet.createRow(rowNum++);
        cell = row.createCell(1);
        cell.setCellValue("외주비 소계");
        sheet.addMergedRegion( new CellRangeAddress(40,40,1,5));
        cell = row.createCell(6);
        cell.setCellValue("=SUM(G32:G40)");

        row = sheet.createRow(rowNum++);
        row = sheet.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("내부 인건비 + 외주비 합계");
        sheet.addMergedRegion( new CellRangeAddress(42,42,0,5));
        cell = row.createCell(6);
        cell.setCellValue("=G29+G41");

        row = sheet.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("기술비 5%");
        sheet.addMergedRegion( new CellRangeAddress(43,43,0,5));
        cell = row.createCell(6);
        cell.setCellValue("=INT(G43*A44/100)");
        
        row = sheet.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("부가가치세 10%");
        sheet.addMergedRegion( new CellRangeAddress(44,44,0,5));
        cell = row.createCell(6);
        cell.setCellValue("=INT((G43+G44)*A45/100)");
        
        row = sheet.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("절사 전 합계");
        sheet.addMergedRegion( new CellRangeAddress(45,45,0,5));
        cell = row.createCell(6);
        cell.setCellValue("=G43+G44+G45");
        
        row = sheet.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("절            사");
        sheet.addMergedRegion( new CellRangeAddress(46,46,0,5));
        cell = row.createCell(6);
        cell.setCellValue("=G48-G46");
        
        row = sheet.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("총    계 (부가세 포함)");
        sheet.addMergedRegion( new CellRangeAddress(47,47,0,5));
        cell = row.createCell(6);
        cell.setCellValue("=ROUNDDOWN(G46,-4)");
        
        
        // 컨텐츠 타입과 파일명 지정
        response.setContentType("ms-vnd/excel");
        //response.setHeader("Content-Disposition", "attachment;filename=example.xls");
        response.setHeader("Content-Disposition", "attachment;filename=TEST.xlsx");

        // Excel File Output
        wb.write(response.getOutputStream());
        
    }
}
