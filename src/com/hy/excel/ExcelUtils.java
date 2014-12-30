package com.hy.excel;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;

public class ExcelUtils {

	private static ExcelUtils excelUtils = new ExcelUtils();

	/**
	 * 生成Excel，输出到 {@link OutputStream}
	 * 
	 * @param title
	 *            标题
	 * @param headerList
	 *            列字段名
	 * @param valueList
	 *            内容
	 * @param os
	 *            {@link OutputStream}
	 * @return
	 */
	public static boolean exportExcel(String title, List<String> headerList,
			List<List<String>> valueList, OutputStream os) {

		return exportExcel(title, headerList, valueList, 50000, 50000, os);
	}

	/**
	 * 生成Excel，输出到 {@link OutputStream}
	 * 
	 * @param title
	 *            标题
	 * @param headerList
	 *            列字段名
	 * @param valueList
	 *            内容
	 * @param maxRow
	 *            最大行数
	 * @param maxCell
	 *            最大列数
	 * @param os
	 *            {@link OutputStream}
	 * @return
	 */
	public static boolean exportExcel(String title, List<String> headerList,
			List<List<String>> valueList, int maxRow, int maxCell,
			OutputStream os) {
		boolean resultBool = true;
		int sheetIndex = 0;
		if (title == null || "".equals(title) || headerList == null
				|| headerList.size() == 0 || valueList == null
				|| valueList.size() == 0 || os == null) {
			resultBool = false;
			return resultBool;
		}
		maxCell = maxCell > 50000 ? 50000 : maxCell;
		maxRow = maxRow > 50000 ? 50000 : maxRow;
		maxCell = maxCell <= 0 ? 50000 : maxCell;
		maxRow = maxRow <= 0 ? 50000 : maxRow;
		// 生成工作簿
		HSSFWorkbook wb = new HSSFWorkbook();

		// 根据内容大小生成工作表
		HSSFSheet sheet = wb.createSheet(title);

		// 默认列宽
		sheet.setDefaultColumnWidth(20);
		// 默认行高
		sheet.setDefaultRowHeightInPoints(15);

		// 样式
		HSSFCellStyle titleStyle = wb.createCellStyle();
		HSSFCellStyle headerStyle = wb.createCellStyle();
		HSSFCellStyle contentStyle = wb.createCellStyle();
		HSSFCellStyle dateStyle = wb.createCellStyle();
		// 左边框
		titleStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		headerStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		contentStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		dateStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		// 上边框
		titleStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		headerStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		contentStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		dateStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		// 右边框
		titleStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		headerStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		contentStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		dateStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		// 下边框
		titleStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		headerStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		contentStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		dateStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		// 水平位置
		titleStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		headerStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		contentStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		dateStyle.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
		// 垂直位置
		titleStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		headerStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		contentStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		dateStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);

		// 字体样式
		HSSFFont titleFont = wb.createFont();
		HSSFFont headerFont = wb.createFont();
		HSSFFont contentFont = wb.createFont();
		HSSFFont dateFont = wb.createFont();

		// 字体
		titleFont.setFontName("微软雅黑");
		headerFont.setFontName("微软雅黑");
		contentFont.setFontName("楷体");
		dateFont.setFontName("黑体");

		// 字体大小
		titleFont.setFontHeightInPoints((short) 16);
		headerFont.setFontHeightInPoints((short) 12);
		contentFont.setFontHeightInPoints((short) 11);
		dateFont.setFontHeightInPoints((short) 10);

		// 特殊字体样式
		titleFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		headerFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);

		// 添加字体样式到行样式中
		titleStyle.setFont(titleFont);
		headerStyle.setFont(headerFont);
		contentStyle.setFont(contentFont);

		// 合并表格
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0,
				(headerList.size() > maxCell ? maxCell : headerList.size()) - 1));
		sheet.addMergedRegion(new CellRangeAddress(1, 1, 0,
				(headerList.size() > maxCell ? maxCell : headerList.size()) - 1));

		// 创建标题-行(第一行，下标为0)
		Row titleRow = sheet.createRow(0);
		titleRow.setHeightInPoints(25);

		// 创建标题-标题内容列
		Cell titleCell = titleRow.createCell(0);
		titleCell.setCellStyle(titleStyle);
		titleCell.setCellValue(title);

		// 创建标题-标题空白列
		for (int i = 1; i < (headerList.size() > maxCell ? maxCell : headerList
				.size()); i++) {
			titleCell = titleRow.createCell(i);
			titleCell.setCellStyle(titleStyle);
		}

		// 创建制表日期
		SimpleDateFormat format = new SimpleDateFormat("yyyy年MM月dd日 HH:mm:ss");

		// 创建制表日期-行(第二行，下标为1)
		Row dateRow = sheet.createRow(1);

		// 创建制表日期-内容列
		Cell dateCell = dateRow.createCell(0);
		dateCell.setCellStyle(dateStyle);
		dateCell.setCellValue("制表日期："
				+ format.format(new Date(System.currentTimeMillis())));

		// 创建制表日期-空白列
		for (int i = 1; i < (headerList.size() > maxCell ? maxCell : headerList
				.size()); i++) {
			dateCell = dateRow.createCell(i);
			dateCell.setCellStyle(dateStyle);
		}

		// 创建表头 -行(第三行，下标为2)
		Row headerRow = sheet.createRow(2);
		for (int i = 0; i < (headerList.size() > maxCell ? maxCell : headerList
				.size()); i++) {
			// 创建表头-列
			Cell headerCell = headerRow.createCell(i);
			headerCell.setCellStyle(headerStyle);
			headerCell.setCellValue(headerList.get(i));
		}

		// 创建内容行(第四行开始，下标为3)
		int i = 0;
		int maxRowIndex = (valueList.size() > maxRow ? maxRow : valueList
				.size());
		while (i < maxRowIndex) {
			Row contentRow = sheet.createRow(3 + i);
			int maxCellNum = valueList.get(0).size() > maxCell ? maxCell
					: valueList.get(0).size();
			maxCellNum = maxCellNum > headerList.size() ? headerList.size()
					: maxCellNum;

			// 创建内容列-内容列
			for (int cellIndex = 0; cellIndex < maxCellNum; cellIndex++) {
				Cell contentCell = contentRow.createCell(cellIndex);
				contentCell.setCellStyle(contentStyle);
				contentCell.setCellValue(valueList.get(0).get(cellIndex));
			}

			// 创建内容列-空白列
			for (int cellIndex = maxCellNum; cellIndex < (headerList.size() > maxCell ? maxCell
					: headerList.size()); cellIndex++) {
				Cell contentCell = contentRow.createCell(cellIndex);
				contentCell.setCellStyle(contentStyle);
			}

			valueList.remove(0);
			i++;
		}

		if (valueList.size() > 0) {
			resultBool = excelUtils.exportExcelNextSheet(wb, sheetIndex, title,
					headerList, valueList, maxRow, maxCell);
		}

		if (resultBool) {
			try {
				wb.write(os);
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
				resultBool = false;
			} finally {
				try {
					if (wb != null)
						wb.close();
					if(os!=null)
						os.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
					resultBool = false;
				}
			}
		}

		return resultBool;
	}

	/**
	 * 内容大于maxRow时，创建新的Sheet
	 * 
	 * @param wb
	 *            {@link HSSFWorkbook}
	 * @param sheetIndex
	 *            区分sheet名称用
	 * @param title
	 *            标题
	 * @param headerList
	 *            列字段名
	 * @param valueList
	 *            内容
	 * @param maxRow
	 *            最大行数
	 * @param maxCell
	 *            最大列数
	 * @return
	 */
	private boolean exportExcelNextSheet(HSSFWorkbook wb, int sheetIndex,
			String title, List<String> headerList,
			List<List<String>> valueList, int maxRow, int maxCell) {
		boolean resultBool = true;
		int currsheetIndex = sheetIndex + 1;
		if (currsheetIndex >= 255) {
			resultBool = false;
			return resultBool;
		}
		// 根据内容大小生成工作表
		HSSFSheet sheet = wb.createSheet(title + " - " + currsheetIndex);

		// 默认列宽
		sheet.setDefaultColumnWidth(20);
		// 默认行高
		sheet.setDefaultRowHeightInPoints(15);

		// 样式
		HSSFCellStyle titleStyle = wb.createCellStyle();
		HSSFCellStyle headerStyle = wb.createCellStyle();
		HSSFCellStyle contentStyle = wb.createCellStyle();
		HSSFCellStyle dateStyle = wb.createCellStyle();
		// 左边框
		titleStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		headerStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		contentStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		dateStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		// 上边框
		titleStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		headerStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		contentStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		dateStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		// 右边框
		titleStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		headerStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		contentStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		dateStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		// 下边框
		titleStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		headerStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		contentStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		dateStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		// 水平位置
		titleStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		headerStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		contentStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		dateStyle.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
		// 垂直位置
		titleStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		headerStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		contentStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		dateStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);

		// 字体样式
		HSSFFont titleFont = wb.createFont();
		HSSFFont headerFont = wb.createFont();
		HSSFFont contentFont = wb.createFont();
		HSSFFont dateFont = wb.createFont();

		// 字体
		titleFont.setFontName("微软雅黑");
		headerFont.setFontName("微软雅黑");
		contentFont.setFontName("楷体");
		dateFont.setFontName("黑体");

		// 字体大小
		titleFont.setFontHeightInPoints((short) 16);
		headerFont.setFontHeightInPoints((short) 12);
		contentFont.setFontHeightInPoints((short) 11);
		dateFont.setFontHeightInPoints((short) 10);

		// 特殊字体样式
		titleFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		headerFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);

		// 添加字体样式到行样式中
		titleStyle.setFont(titleFont);
		headerStyle.setFont(headerFont);
		contentStyle.setFont(contentFont);

		// 合并表格
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0,
				(headerList.size() > maxCell ? maxCell : headerList.size()) - 1));
		sheet.addMergedRegion(new CellRangeAddress(1, 1, 0,
				(headerList.size() > maxCell ? maxCell : headerList.size()) - 1));

		// 创建标题-行(第一行，下标为0)
		Row titleRow = sheet.createRow(0);
		titleRow.setHeightInPoints(25);

		// 创建标题-标题内容列
		Cell titleCell = titleRow.createCell(0);
		titleCell.setCellStyle(titleStyle);
		titleCell.setCellValue(title);

		// 创建标题-标题空白列
		for (int i = 1; i < (headerList.size() > maxCell ? maxCell : headerList
				.size()); i++) {
			titleCell = titleRow.createCell(i);
			titleCell.setCellStyle(titleStyle);
		}

		// 创建制表日期
		SimpleDateFormat format = new SimpleDateFormat("yyyy年MM月dd日 HH:mm:ss");

		// 创建制表日期-行(第二行，下标为1)
		Row dateRow = sheet.createRow(1);

		// 创建制表日期-内容列
		Cell dateCell = dateRow.createCell(0);
		dateCell.setCellStyle(dateStyle);
		dateCell.setCellValue("制表日期："
				+ format.format(new Date(System.currentTimeMillis())));

		// 创建制表日期-空白列
		for (int i = 1; i < (headerList.size() > maxCell ? maxCell : headerList
				.size()); i++) {
			dateCell = dateRow.createCell(i);
			dateCell.setCellStyle(dateStyle);
		}

		// 创建表头 -行(第三行，下标为2)
		Row headerRow = sheet.createRow(2);
		for (int i = 0; i < (headerList.size() > maxCell ? maxCell : headerList
				.size()); i++) {
			// 创建表头-列
			Cell headerCell = headerRow.createCell(i);
			headerCell.setCellStyle(headerStyle);
			headerCell.setCellValue(headerList.get(i));
		}

		// 创建内容行(第四行开始，下标为3)
		int i = 0;
		int maxRowIndex = (valueList.size() > maxRow ? maxRow : valueList
				.size());
		while (i < maxRowIndex) {
			Row contentRow = sheet.createRow(3 + i);
			int maxCellNum = valueList.get(0).size() > maxCell ? maxCell
					: valueList.get(0).size();
			maxCellNum = maxCellNum > headerList.size() ? headerList.size()
					: maxCellNum;

			// 创建内容列-内容列
			for (int cellIndex = 0; cellIndex < maxCellNum; cellIndex++) {
				Cell contentCell = contentRow.createCell(cellIndex);
				contentCell.setCellStyle(contentStyle);
				contentCell.setCellValue(valueList.get(0).get(cellIndex));
			}

			// 创建内容列-空白列
			for (int cellIndex = maxCellNum; cellIndex < (headerList.size() > maxCell ? maxCell
					: headerList.size()); cellIndex++) {
				Cell contentCell = contentRow.createCell(cellIndex);
				contentCell.setCellStyle(contentStyle);
			}

			valueList.remove(0);
			i++;
		}

		if (valueList.size() > 0) {
			resultBool = excelUtils.exportExcelNextSheet(wb, currsheetIndex,
					title, headerList, valueList, maxRow, maxCell);
		}

		return resultBool;
	}

	/**
	 * 读取Excel(xls)文件的内容
	 * @param is {@link InputStream} 输入流
	 * @param hasHeader boolean 是否有标题(只限第一行)
	 * @return {@link List}<{@link List}<{@link String}>> 所有行<所有列>
	 */
	public static List<List<String>> importExcel(InputStream is,int maxCell,boolean hasHeader) {
		List<List<String>> resultList = null;
		Workbook wb = null;
		try {
			wb = WorkbookFactory.create(is);
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		Sheet sheet = wb!=null?wb.getSheetAt(0):null;
		if(sheet!=null) {
			int maxCellNum = maxCell>50000?50000:maxCell; 
			if(maxCellNum>0) {
				SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
				DecimalFormat numberFormat = new DecimalFormat("0.00");
				resultList = new LinkedList<List<String>>();
				for(Row row : sheet) {
					List<String> lineList = new LinkedList<String>();
					if(row.getRowNum()==0) {
						if(hasHeader)
							continue;
					}
					for(int i=0;i<maxCellNum;i++) {
						Cell cell = row.getCell(i);
						if(cell==null){
							lineList.add("");
							continue;
						}
						switch(cell.getCellType()) {
							case Cell.CELL_TYPE_BLANK:
								lineList.add("");
								break;
							case Cell.CELL_TYPE_BOOLEAN:
								lineList.add(Boolean.toString(cell.getBooleanCellValue()));
								break;
							case Cell.CELL_TYPE_ERROR:
								lineList.add("");
								break;
							case Cell.CELL_TYPE_FORMULA:
								lineList.add("");
								break;
							case Cell.CELL_TYPE_NUMERIC:
								if(HSSFDateUtil.isCellDateFormatted(cell)) {
									lineList.add(dateFormat.format(cell.getDateCellValue()));
								} else {
									lineList.add(numberFormat.format(cell.getNumericCellValue()));
								}
								break;
							case Cell.CELL_TYPE_STRING:
								lineList.add(cell.getStringCellValue());
								break;
						}
					}
					resultList.add(lineList);
				}
			}
		}
		try {
			if(wb!=null)
				wb.close();
			if(is!=null)
				is.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return resultList;
	}

}
