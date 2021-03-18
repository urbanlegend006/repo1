package utility;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

public class excelReader {

	
	

	private XSSFWorkbook getWorkBook(String filename) {

		try {
			FileInputStream fis = new FileInputStream(filename);
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			return workbook;
		} catch (IOException e) {
			e.printStackTrace();
		}
		return null;
	}

	private XSSFSheet getSheet(String filename) {
		XSSFWorkbook workbook = getWorkBook(filename);
		if (workbook != null) {
			XSSFSheet sheet = workbook.getSheetAt(0);
			return sheet;
		}
		return null;
	}

	private List<List<String>> getRowList(String filename) {

		XSSFSheet sheet = getSheet(filename);

		List<List<String>> rowList = new LinkedList<List<String>>();
		
		int totalRows = sheet.getLastRowNum();
		int headerRow = headerRowNum(sheet);
		//System.out.println("First Row Count\nFrom My function :"+headerRow+" ,  nd by XSSF :"+sheet.getFirstRowNum());
		Row row;

		for (int i = headerRow + 1; i <= totalRows; i++) {
			row = sheet.getRow(i);
			if (getColumnList(row) != null) {
				rowList.add(getColumnList(row));
			}
		}

		return rowList;
	}

	private List<String> getColumnList(Row row) {

		List<String> columnList = null;
		columnList = new ArrayList<String>();

		for (int j = 0; j < row.getLastCellNum(); j++) {

			Cell cell = row.getCell(j);
	
			if (cell.getCellType() == CellType.STRING) {
				columnList.add(cell.getStringCellValue());
			} else if (cell.getCellType() == CellType.NUMERIC) {
				columnList.add("'" + cell.getNumericCellValue());
			} else {
				columnList.add("");
			}
		}
		return columnList;
	}

	private int headerRowNum(XSSFSheet sheet) {
		int totalRows = sheet.getLastRowNum();
		int headerRow = 0;

		Row row;
		for (int i = 0; i < totalRows; i++) {
			int flag = 0;
			row = sheet.getRow(i);
			int lastColumn = row.getLastCellNum();
			for (int k = 0; k < lastColumn; k++) {
				if (row.getCell(k).getStringCellValue() == "") {
					flag = 1;
					break;
				}

			}
			if (flag == 0) {
				headerRow = i;
			}

		}
		return (headerRow-1);
	}

	public List<List<String>> getExcelDataAsLists(String filename) {

		return getRowList(filename);

	}
	
	

}
