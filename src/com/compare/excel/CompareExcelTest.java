/**
 * 
 */
package com.compare.excel;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * @author akhileshP
 *
 */
public class CompareExcelTest {

	public static void main(String[] args) {
		try (FileInputStream excellFile1 = new FileInputStream(new File("FILE_PATH1"));
				FileInputStream excellFile2 = new FileInputStream(new File("FILE_PATH2"));

				HSSFWorkbook workbook1 = new HSSFWorkbook(excellFile1);
				HSSFWorkbook workbook2 = new HSSFWorkbook(excellFile2);) {

			HSSFSheet sheet1 = workbook1.getSheetAt(0);
			HSSFSheet sheet2 = workbook2.getSheetAt(0);

			if (compareSheets(sheet1, sheet2)) {
				System.out.println("\n\nThe two excel sheets are Equal");
			} else {
				System.out.println("\n\nThe two excel sheets are Not Equal");
			}

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public static boolean compareSheets(HSSFSheet sheet1, HSSFSheet sheet2) {
		int firstRow1 = sheet1.getFirstRowNum();
		int lastRow1 = sheet1.getLastRowNum();
		boolean equalSheets = true;
		for (int i = firstRow1; i <= lastRow1; i++) {
			System.out.println("\n\nComparing Row " + i);
			HSSFRow row1 = sheet1.getRow(i);
			HSSFRow row2 = sheet2.getRow(i);
			if (!compareRows(row1, row2)) {
				equalSheets = false;
				System.out.println("Row " + i + " - Not Equal");
			}
		}
		return equalSheets;
	}

	public static boolean compareRows(HSSFRow row1, HSSFRow row2) {
		if ((row1 == null) && (row2 == null)) {
			return true;
		} else if ((row1 == null) || (row2 == null)) {
			return false;
		}

		int firstCell1 = row1.getFirstCellNum();
		int lastCell1 = row1.getLastCellNum();
		boolean equalRows = true;

		for (int i = firstCell1; i <= lastCell1; i++) {
			HSSFCell cell1 = row1.getCell(i);
			HSSFCell cell2 = row2.getCell(i);
			if (!compareCells(cell1, cell2)) {
				equalRows = false;
				System.err.println("       Cell " + i + " - Not Equal" + "; Value of Cell " + i + " is \"" + cell1
						+ "\" - Value of Cell " + i + " is \"" + cell2 + "\"");
			}
		}
		return equalRows;
	}

	public static boolean compareCells(HSSFCell cell1, HSSFCell cell2) {
		if ((cell1 == null) && (cell2 == null)) {
			return true;
		} else if ((cell1 == null) || (cell2 == null)) {
			return false;
		}

		boolean equalCells = false;
		int type1 = cell1.getCellType();
		int type2 = cell2.getCellType();
		if (type1 == type2) {
			if (cell1.getCellStyle().equals(cell2.getCellStyle())) {
				switch (cell1.getCellType()) {
				case HSSFCell.CELL_TYPE_FORMULA:
					if (cell1.getCellFormula().equals(cell2.getCellFormula())) {
						equalCells = true;
					}
					break;
				case HSSFCell.CELL_TYPE_NUMERIC:
					if (cell1.getNumericCellValue() == cell2.getNumericCellValue()) {
						equalCells = true;
					}
					break;
				case HSSFCell.CELL_TYPE_STRING:
					if (cell1.getStringCellValue().equals(cell2.getStringCellValue())) {
						equalCells = true;
					}
					break;
				case HSSFCell.CELL_TYPE_BLANK:
					if (cell2.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
						equalCells = true;
					}
					break;
				case HSSFCell.CELL_TYPE_BOOLEAN:
					if (cell1.getBooleanCellValue() == cell2.getBooleanCellValue()) {
						equalCells = true;
					}
					break;
				case HSSFCell.CELL_TYPE_ERROR:
					if (cell1.getErrorCellValue() == cell2.getErrorCellValue()) {
						equalCells = true;
					}
					break;
				default:
					if (cell1.getStringCellValue().equals(cell2.getStringCellValue())) {
						equalCells = true;
					}
					break;
				}
			} else {
				return false;
			}
		} else {
			return false;
		}
		return equalCells;
	}
}
