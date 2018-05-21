package com.pradeipk.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.pradeipk.excel.util.Util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

public class ReadExcel {

	private static final String FILE_NAME = "F:\\3_Utkarsh\\Finance\\utkarsh.xlsx";
	private static final String FILE_NAME_O = "F:/3_Utkarsh/Finance/uts.xlsx";

	public static void main(String[] args) {

		try {

			FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
			FileOutputStream output = new FileOutputStream(new File(FILE_NAME_O));
			//Workbook oworkbook = new XSSFWorkbook();
			Workbook workbook = new XSSFWorkbook(excelFile);
			Sheet receiptByDate = workbook.getSheet("ReceiptByDate");

			System.out.println("Sheet Name" + " : " + receiptByDate.getSheetName());

			Iterator<Row> iterator = receiptByDate.iterator();

			while (iterator.hasNext()) {

				Row currentRow = iterator.next();
				int row_num = currentRow.getRowNum();

				// if(row_num==0) return;
				System.out.println("\n--Row Number " + row_num);
				Iterator<Cell> cellIterator = currentRow.iterator();
				StringBuilder memcode = null;
				int index;
				while (cellIterator.hasNext()) {
					if (row_num == 0)
						break;
					Cell currentCell = cellIterator.next();
					index = currentCell.getColumnIndex();
					// getCellTypeEnum shown as deprecated for version 3.15
					// getCellTypeEnum ill be renamed to getCellType starting from version 4.0
					if (index == 1) {
						memcode = new StringBuilder();
						if (currentCell.getCellTypeEnum().toString().equalsIgnoreCase("NUMERIC")) {
							memcode = memcode.append(Util.getDate(currentCell.getDateCellValue()) + "-");
						} else {
							memcode = new StringBuilder(currentCell.getStringCellValue() + "-");
						}

					} else if (index == 4) {
						// Name
						memcode.append(currentCell.getStringCellValue().substring(0, 3));
						// System.out.print(currentCell.getNumericCellValue() + "--");

					} else if (index == 7) {
						// Village
						memcode.append(currentCell.getStringCellValue().substring(0, 3));
						// System.out.print(currentCell.getNumericCellValue() + "--");
						System.out.println(memcode);
						currentRow.createCell(2).setCellType(CellType.STRING);
						currentRow.getCell(2).setCellValue(memcode.toString());
						// currentRow.getCell(2).setCellValue(memcode.toString());
						break;
					}

				}
				if (row_num == 10)
					break;
			}
			workbook.write(output);
		workbook.close();
			System.out.println();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}
}
