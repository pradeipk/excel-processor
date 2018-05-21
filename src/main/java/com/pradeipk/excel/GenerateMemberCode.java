package com.pradeipk.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.pradeipk.excel.util.Util;

public class GenerateMemberCode {

	private final String FILE_NAME = "Master_Sheet.xlsx";
	private final String FILE_PATH_READ = "";
	private final String FILE_PATH_WRITE = "E:/Utkarsh/";
	private final String SHEET_NAME = "ReceiptByDate";
	private final int ROWS_TO_PROCESS = 456;

	public void generateMemberCode() throws IOException {

		Workbook workbook = getWorkBook();
		FileOutputStream output = new FileOutputStream(new File(FILE_PATH_WRITE
				+ FILE_NAME));
		if (workbook == null)
			return;

		Sheet receiptByDate = workbook.getSheet(SHEET_NAME);
		System.out.println("Sheet Name" + " : " + receiptByDate.getSheetName());
		Iterator<Row> rowIterator = receiptByDate.iterator();

		while (rowIterator.hasNext()) {
			Row currentRow = rowIterator.next();
			/*if (currentRow.getRowNum() > ROWS_TO_PROCESS)
				return;*/
			updateRow(currentRow);

		}
		workbook.write(output);
		workbook.close();
	}

	private Workbook getWorkBook() {
		Workbook workbook = null;
		try {
			ClassLoader classLoader = getClass().getClassLoader();
			File file = new File(classLoader.getResource(FILE_NAME).getFile());
			FileInputStream excelFile = new FileInputStream(file);
			workbook = new XSSFWorkbook(excelFile);
			System.out.println(workbook.getNumberOfNames());
		}
		catch (Exception e) {
			System.out.println("Error :" + e.getMessage());
		}

		return workbook;

	}

	private void updateRow(Row row) {

		int row_num = row.getRowNum();
		Iterator<Cell> cellIterator = row.iterator();
		StringBuilder memcode = new StringBuilder();

		while (cellIterator.hasNext()) {
			if (row_num == 0)
				break;

			Cell currentCell = cellIterator.next();
			formCode(currentCell, memcode);
		}
		System.out.println(memcode);
		row.createCell(2).setCellType(CellType.STRING);
		row.getCell(2).setCellValue(memcode.toString());

	}

	private String getValue(Cell cell) {
		System.out.println(cell.getCellTypeEnum());

		switch (cell.getCellTypeEnum()) {
		case BOOLEAN:
			return "true";
		case NUMERIC:
			return Util.getDate(cell.getDateCellValue());
		default:
			return cell.getStringCellValue();
		}
	}

	private void formCode(Cell cell, StringBuilder memcode) {
		int index = cell.getColumnIndex();
		if (isValidIndex(index)) {
			memcode.append(getValue(cell));
			memcode.append("-");
		}

	}

	private boolean isValidIndex(int index) {
		boolean flag = false;

		if (index == 1 || index == 4 || index == 7) {
			flag = true;
		}
		return flag;
	}

}
