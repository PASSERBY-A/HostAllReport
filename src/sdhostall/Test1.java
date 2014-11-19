package sdhostall;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
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





public class Test1 {
	
	public List<Map<String,String>> queryData(String beginDate,String endDate){
		
		List<Map<String,String>> result = new ArrayList<Map<String,String>>();
		
		Connection connection =  DBManager.getInstance().getConnection();
		
		Statement stmt = null;
		
		ResultSet rs = null;
		StringBuffer sql = new StringBuffer();
		
		sql.append("select caption,host,hardinfo,ip,bizsystem,cpuusage, memusage,swapusage,fileusage,diskusage,procnum,livetime,NTP_State,alarm1,alarm2, alarm3,alarm4, alarm5");
		
		sql.append(" from SUMMARY_REPORT_DATA where kpi_time ");
		
		sql.append("between to_date('"+beginDate+"', 'YYYY-MM-DD') and   to_date('"+endDate+"', 'YYYY-MM-DD')");
		
		try 
		{
			
		 stmt = connection.createStatement();
			
		 rs = stmt.executeQuery(sql.toString());
		
		while(rs.next())
		{
			Map<String,String> map = new HashMap<String, String>();
			map.put("cell0",rs.getString(1));
			map.put("cell1",rs.getString(2));
			map.put("cell2",rs.getString(3));
			map.put("cell3",rs.getString(4));
			map.put("cell4",rs.getString(5));
			map.put("cell5",rs.getString(6));
			map.put("cell6",rs.getString(7));
			map.put("cell7",rs.getString(8));
			map.put("cell8",rs.getString(9));
			map.put("cell9",rs.getString(10));
			map.put("cell10",rs.getString(11));
			map.put("cell11",rs.getString(12));
			map.put("cell12",rs.getString(13));
			map.put("cell13",rs.getString(14));
			map.put("cell14",rs.getString(15));
			map.put("cell15",rs.getString(16));
			map.put("cell16",rs.getString(17));
			map.put("cell17",rs.getString(18));
			result.add(map);
		}
		
			
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally{
			
			try {
				DBManager.getInstance().release(connection, stmt, rs);
			} catch (SQLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
		}
		
		return result;
		
	}
	
	
	
	
	public static void main(String[] args) throws Exception {

		Test1 g = new Test1();
		
		HSSFWorkbook wb = g.generate("2014-11-12", "2014-11-12");
		
		//���ﶨ�����Լ���Ҫ������� 
		FileOutputStream os = new FileOutputStream("c:\\a.xls");
		
		wb.write(os);
		
		os.flush();
		os.close();
		
	}


	public HSSFWorkbook generate(String begin,String end) throws IOException
	{
		Test1 g = new Test1();
		
		
		
		// ����Excel�Ĺ������ Workbook,��Ӧ��һ��excel�ĵ�
		HSSFWorkbook wb = new HSSFWorkbook();
		// ����Excel�Ĺ���sheet,��Ӧ��һ��excel�ĵ���tab
		HSSFSheet sheet = wb.createSheet("sheet1");
		// ����excelÿ�п��
		sheet.setColumnWidth(0, 4000);
		sheet.setColumnWidth(1, 3500);
		
		HSSFFont font = setFont(wb);
		
		HSSFCellStyle style = setStyle(wb, font);
		
		createTitle(sheet,style);
		
		List<Map<String,String>> data =  g.queryData(begin,end);
		
		
		HSSFRow contentRow = sheet.createRow(1);
		// ���õ�Ԫ�����ʽ��ʽ
		HSSFCellStyle style1 = wb.createCellStyle();
		for(int i=0;i<data.size();i++)
		{
			
			// ���õ�Ԫ�����ݸ�ʽ
			 style1 = wb.createCellStyle();
			//style1.setDataFormat(HSSFDataFormat.getBuiltinFormat("h:mm:ss"));
			style1.setWrapText(true);// �Զ�����
		    contentRow = sheet.createRow(i+1);
		    style1.setVerticalAlignment((short)1);
			
			
			for(int j=0;j<data.get(i).size();j++)
			{
				// ���õ�Ԫ�����ʽ��ʽ
				HSSFCell cell = contentRow.createCell(j);
				cell.setCellStyle(style1);
				String content = data.get(i).get("cell"+String.valueOf(j));
				
				
				if((j==9 ||j==8||j==3 )&&!"".equals(content) && null!=content)
				{
					String []c = content.split(",");
					StringBuffer sb = new StringBuffer();
					for(int ii=0;ii<c.length;ii++)
					{
						sb.append(c[ii]+"\n");
						
						
					}
					content = sb.toString();
				}
				cell.setCellValue(content);
			}
			
			
		}/**/
		
		
		return wb;
		
		
		
	}
	
	
	

	private static void createTitle(HSSFSheet sheet, HSSFCellStyle style) {
		
		// ����Excel��sheet��һ��
		HSSFRow row = sheet.createRow(0);
		
		row.setHeight((short) 500);// �趨�еĸ߶�
		
		// ����һ��Excel�ĵ�Ԫ��
		HSSFCell cell0 = row.createCell(0);
		// �ϲ���Ԫ��(startRow��endRow��startColumn��endColumn)
		//sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 2));
		// ��Excel�ĵ�Ԫ��������ʽ�͸�ֵ
		cell0.setCellStyle(style);
		cell0.setCellValue("ҵ��ϵͳ");
		
		HSSFCell cell1 = row.createCell(1);
		cell1.setCellStyle(style);
		cell1.setCellValue("ϵͳ����");
		
		
		HSSFCell cell2 = row.createCell(2);
		cell2.setCellStyle(style);
		cell2.setCellValue("������");
		
		
		HSSFCell cell3 = row.createCell(3);
		cell3.setCellStyle(style);
		cell3.setCellValue("Ӳ����Ϣ");
		
		
		HSSFCell cell4 = row.createCell(4);
		cell4.setCellStyle(style);
		cell4.setCellValue("IP��ַ");
		
		
		HSSFCell cell5 = row.createCell(5);
		cell5.setCellStyle(style);
		cell5.setCellValue("CPUʹ����");
		//HSSFCellStyle style5 = wb.createCellStyle();
		//cell5.set
		
		HSSFCell cell6 = row.createCell(6);
		cell6.setCellStyle(style);
		cell6.setCellValue("�ڴ�ʹ����");
		
		HSSFCell cell7 = row.createCell(7);
		cell7.setCellStyle(style);
		cell7.setCellValue("SWAPʹ����");
		
		
		HSSFCell cell8 = row.createCell(8);
		cell8.setCellStyle(style);
		cell8.setCellValue("�ļ�ʹ����");
		
		
		HSSFCell cell9 = row.createCell(9);
		cell9.setCellStyle(style);
		cell9.setCellValue("I/Oʹ����");
		
		
		HSSFCell cell10 = row.createCell(10);
		cell10.setCellStyle(style);
		cell10.setCellValue("������");
		
		
		HSSFCell cell11 = row.createCell(11);
		cell11.setCellStyle(style);
		cell11.setCellValue("����ʱ��");
		
		
		HSSFCell cell12 = row.createCell(12);
		cell12.setCellStyle(style);
		cell12.setCellValue("NTP����״̬");
		
		
		HSSFCell cell13 = row.createCell(13);
		cell13.setCellStyle(style);
		cell13.setCellValue("���ظ澯");
		
		
		HSSFCell cell14 = row.createCell(14);
		cell14.setCellStyle(style);
		cell14.setCellValue("��Ҫ�澯");
		
		HSSFCell cell15 = row.createCell(15);
		cell15.setCellStyle(style);
		cell15.setCellValue("��Ҫ�澯");
		
		HSSFCell cell16 = row.createCell(16);
		cell16.setCellStyle(style);
		cell16.setCellValue("һ��澯");
		
		
		HSSFCell cell17 = row.createCell(17);
		cell17.setCellStyle(style);
		cell17.setCellValue("��Ϣ");
		
		
		
	}



	private static HSSFFont setFont(HSSFWorkbook wb) {
		// ����������ʽ
		HSSFFont font = wb.createFont();
		font.setFontName("SansSerif");
		font.setBoldweight((short) 100);
		font.setFontHeight((short) 200);
		//font.setColor(HSSFColor.GREY_25_PERCENT.i);
		return font;
	}
	
	private static HSSFCellStyle setStyle(HSSFWorkbook wb, HSSFFont font) {
		// ������Ԫ����ʽ
		HSSFCellStyle style = wb.createCellStyle();
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		//style.setFillForegroundColor(HSSFColor.LIGHT_TURQUOISE.index);
		style.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
		style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		// ���ñ߿�
		style.setBottomBorderColor(HSSFColor.BLACK.index);
		style.setTopBorderColor(HSSFColor.BLACK.index);
		style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style.setBorderRight(HSSFCellStyle.BORDER_THIN);
		style.setBorderTop(HSSFCellStyle.BORDER_THIN);
		style.setFont(font);// ��������
		return style;
	}
}
