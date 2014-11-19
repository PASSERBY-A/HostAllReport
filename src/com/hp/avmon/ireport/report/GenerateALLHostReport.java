package com.hp.avmon.ireport.report;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.springframework.context.ApplicationContext;
import org.springframework.context.support.ClassPathXmlApplicationContext;
import org.springframework.jdbc.core.JdbcTemplate;

public class GenerateALLHostReport {

	@SuppressWarnings("unchecked")
	public List<Map<String, String>> queryData(JdbcTemplate jdbc,String beginDate, String endDate) {

		List<Map<String, String>> result = new ArrayList<Map<String, String>>();
		List<Map<String, String>> list = new ArrayList<Map<String, String>>();

//		ResultSet rs = null;
		StringBuffer sql = new StringBuffer();

		sql.append("select caption as \"caption\",host as \"host\",hardinfo as \"hardinfo\",ip as \"ip\",bizsystem as \"bizsystem\",cpuusage as \"cpuusage\", memusage as \"memusage\",swapusage as \"swapusage\","
				+ "fileusage as \"fileusage\",diskusage as \"diskusage\",procnum as \"procnum\",livetime as \"livetime\",NTP_State as \"ntpstate\",alarm1 as \"alarm1\",alarm2 as \"alarm2\", alarm3 as \"alarm3\",alarm4 as \"alarm4\", alarm5 as \"alarm5\"");

		sql.append(" from SUMMARY_REPORT_DATA where kpi_time ");

		sql.append("between to_date('" + beginDate
				+ "', 'YYYY-MM-DD') and  to_date('" + endDate
				+ "', 'YYYY-MM-DD')");
		list = jdbc.queryForList(sql.toString());
		if(list.size() > 0){ 
			for (Map<String, String> qmap : list) {
				Map<String, String> map = new HashMap<String, String>();
				map.put("cell0", qmap.get("bizsystem")); 
				map.put("cell1", qmap.get("caption"));
				map.put("cell2", qmap.get("host"));
				map.put("cell3", qmap.get("hardinfo"));
				map.put("cell4", qmap.get("ip"));
				map.put("cell5", qmap.get("cpuusage"));
				map.put("cell6", qmap.get("memusage"));
				map.put("cell7", qmap.get("swapusage"));
				map.put("cell8", qmap.get("fileusage"));
				map.put("cell9", qmap.get("diskusage"));
			//	map.put("cell10", qmap.get("procnum"));
			//	map.put("cell11", qmap.get("livetime"));
				map.put("cell10", qmap.get("ntpstate"));
				map.put("cell11", qmap.get("alarm1"));
				map.put("cell12", qmap.get("alarm2"));
				map.put("cell13", qmap.get("alarm3"));
				map.put("cell14", qmap.get("alarm4"));
				map.put("cell15", qmap.get("alarm5"));
				result.add(map);
			} 
		}
		
		return result;

	}

//	public static void main(String[] args) throws Exception {
//
//		GenerateALLHostReport g = new GenerateALLHostReport();
//
//		HSSFWorkbook wb = g.generate("2014-11-12", "2014-11-12");
//
//		// 这里定义你自己需要的输出流
//		FileOutputStream os = new FileOutputStream("c:\\a.xls");
//
//		wb.write(os);
//
//		os.flush();
//		os.close();
//
//	}
	
/*	public static void main(String[] args) throws IOException {
		
		GenerateALLHostReport g = new GenerateALLHostReport();
		
		ApplicationContext app = new ClassPathXmlApplicationContext("applicationContext.xml");
		
		JdbcTemplate jdbc = app.getBean(JdbcTemplate.class);
		
		HSSFWorkbook wb = g.generate(jdbc, "2014-11-18", "2014-11-18");
		
		FileOutputStream os = new FileOutputStream("c:\\11111111111.xls");
		
		wb.write(os);
		os.flush();
		os.close();
	}
	*/
	
	
	

	public HSSFWorkbook generate(JdbcTemplate jdbc,String begin, String end) throws IOException {
		GenerateALLHostReport g = new GenerateALLHostReport();

		// 创建Excel的工作书册 Workbook,对应到一个excel文档
		HSSFWorkbook wb = new HSSFWorkbook();
		// 创建Excel的工作sheet,对应到一个excel文档的tab
		HSSFSheet sheet = wb.createSheet("sheet1");
		// 设置excel每列宽度
		sheet.setColumnWidth(0, 4000);
		sheet.setColumnWidth(1, 3500);

		HSSFFont font = setFont(wb);

		HSSFCellStyle style = setStyle(wb, font);

		createTitle(sheet, style);
		
			sheet.autoSizeColumn((short)0); //调整第一列宽度
	        sheet.autoSizeColumn((short)1); //调整第二列宽度
	        sheet.autoSizeColumn((short)2); //调整第三列宽度
	        sheet.autoSizeColumn((short)3); //调整第四列宽度
	        sheet.autoSizeColumn((short)4);
	        sheet.autoSizeColumn((short)5);
	        sheet.autoSizeColumn((short)6);
	        sheet.autoSizeColumn((short)7);
	        sheet.autoSizeColumn((short)8);
	        sheet.autoSizeColumn((short)9);
	        sheet.autoSizeColumn((short)10);
	        sheet.autoSizeColumn((short)11);
	        sheet.autoSizeColumn((short)12);
	        sheet.autoSizeColumn((short)13);
	        sheet.autoSizeColumn((short)14);
	        sheet.autoSizeColumn((short)15);
	        
		List<Map<String, String>> data = g.queryData(jdbc,begin, end);

		HSSFRow contentRow = sheet.createRow(1);
		// 设置单元格的样式格式
		HSSFCellStyle style1 = wb.createCellStyle();
		for (int i = 0; i < data.size(); i++) {

			// 设置单元格内容格式
			style1 = wb.createCellStyle();
			// style1.setDataFormat(HSSFDataFormat.getBuiltinFormat("h:mm:ss"));
			style1.setWrapText(true);// 自动换行
			contentRow = sheet.createRow(i + 1);
			style1.setVerticalAlignment((short) 1);

			for (int j = 0; j < data.get(i).size(); j++) {
				// 设置单元格的样式格式
				HSSFCell cell = contentRow.createCell(j);
				cell.setCellStyle(style1);
				String content = data.get(i).get("cell" + String.valueOf(j));

				if ((j == 9 || j == 8 || j == 4) && !"".equals(content)
						&& null != content) {
					String[] c = content.split(",");
					StringBuffer sb = new StringBuffer();
					for (int ii = 0; ii < c.length; ii++) {
						sb.append(c[ii] + "\n");

					}
					content = sb.toString();
				}
				cell.setCellValue(content);
			}

		}/**/

		return wb;

	}

	private static void createTitle(HSSFSheet sheet, HSSFCellStyle style) {

		// 创建Excel的sheet的一行
		HSSFRow row = sheet.createRow(0);

		row.setHeight((short) 500);// 设定行的高度

		// 创建一个Excel的单元格
		HSSFCell cell0 = row.createCell(0);
		// 合并单元格(startRow，endRow，startColumn，endColumn)
		// sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 2));
		// 给Excel的单元格设置样式和赋值
		cell0.setCellStyle(style);
		cell0.setCellValue("业务系统");

		HSSFCell cell1 = row.createCell(1);
		cell1.setCellStyle(style); 
		cell1.setCellValue("系统类型");

		HSSFCell cell2 = row.createCell(2);
		cell2.setCellStyle(style);
		cell2.setCellValue("主机名");

		HSSFCell cell3 = row.createCell(3);
		cell3.setCellStyle(style);
		cell3.setCellValue("硬件信息");

		HSSFCell cell4 = row.createCell(4);  
		cell4.setCellStyle(style);
		cell4.setCellValue("          IP地址          ");

		HSSFCell cell5 = row.createCell(5);
		cell5.setCellStyle(style);
		cell5.setCellValue("CPU使用率");
		// HSSFCellStyle style5 = wb.createCellStyle();
		// cell5.set

		HSSFCell cell6 = row.createCell(6);
		cell6.setCellStyle(style);
		cell6.setCellValue("内存使用率");

		HSSFCell cell7 = row.createCell(7);
		cell7.setCellStyle(style);
		cell7.setCellValue("SWAP使用率");

		HSSFCell cell8 = row.createCell(8);
		cell8.setCellStyle(style);
		cell8.setCellValue("文件使用率");

		HSSFCell cell9 = row.createCell(9);
		cell9.setCellStyle(style);
		cell9.setCellValue("I/O使用率");
/* 
		HSSFCell cell10 = row.createCell(10);
		cell10.setCellStyle(style);
		cell10.setCellValue("进程数");

		HSSFCell cell11 = row.createCell(11);
		cell11.setCellStyle(style);
		cell11.setCellValue("运行时间");*/

		HSSFCell cell10 = row.createCell(10);
		cell10.setCellStyle(style);
		cell10.setCellValue("NTP服务状态");

		HSSFCell cell11 = row.createCell(11);
		cell11.setCellStyle(style);
		cell11.setCellValue("严重告警");

		HSSFCell cell12 = row.createCell(12);
		cell12.setCellStyle(style);
		cell12.setCellValue("主要告警");

		HSSFCell cell13 = row.createCell(13);
		cell13.setCellStyle(style);
		cell13.setCellValue("次要告警");

		HSSFCell cell14 = row.createCell(14);
		cell14.setCellStyle(style);
		cell14.setCellValue("一般告警"); 

		HSSFCell cell15 = row.createCell(15);
		cell15.setCellStyle(style);
		cell15.setCellValue("信息");

	}

	private static HSSFFont setFont(HSSFWorkbook wb) {
		// 创建字体样式
		HSSFFont font = wb.createFont();
		font.setFontName("SansSerif");
		font.setBoldweight((short) 100);
		font.setFontHeight((short) 200);
		// font.setColor(HSSFColor.GREY_25_PERCENT.i);
		return font;
	}

	private static HSSFCellStyle setStyle(HSSFWorkbook wb, HSSFFont font) {
		// 创建单元格样式
		HSSFCellStyle style = wb.createCellStyle();
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		// style.setFillForegroundColor(HSSFColor.LIGHT_TURQUOISE.index);
		style.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
		style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		// 设置边框
		style.setBottomBorderColor(HSSFColor.BLACK.index);
		style.setTopBorderColor(HSSFColor.BLACK.index);
		style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style.setBorderRight(HSSFCellStyle.BORDER_THIN);
		style.setBorderTop(HSSFCellStyle.BORDER_THIN);
		style.setFont(font);// 设置字体
		return style;
	}
}
