/**
 * 
 */
package com.compare.excel;

import java.io.File;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;

/**
 * @author pramati
 *
 */
public class ExcelDataCompare {

	private static final String CELL_DATA_DOES_NOT_MATCH = "Cell Data does not Match ::";

	List<String> listOfDifferences = new ArrayList<>();

	private static class Locator {
		Workbook workbook;
		Sheet sheet;
		Row row;
		Cell cell;
	}

	public static void main(String[] args) throws Exception {

		Workbook wb1 = WorkbookFactory.create(new File("File_Path"));
		Workbook wb2 = WorkbookFactory.create(new File("File_Path"));
		System.out.println(ExcelDataCompare.compare(wb1, wb2));

	}

	public static List<String> compare(Workbook wb1, Workbook wb2) {
		Locator loc1 = new Locator();
		Locator loc2 = new Locator();
		loc1.workbook = wb1;
		loc2.workbook = wb2;

		ExcelDataCompare excelComparator = new ExcelDataCompare();
		excelComparator.compareSheetData(loc1, loc2);

		return excelComparator.listOfDifferences;
	}

	private void compareDataInAllSheets(Locator loc1, Locator loc2) {
		for (int i = 0; i < loc1.workbook.getNumberOfSheets(); i++) {
			if (loc2.workbook.getNumberOfSheets() <= i)
				return;

			loc1.sheet = loc1.workbook.getSheetAt(i);
			loc2.sheet = loc2.workbook.getSheetAt(i);

			compareDataInSheet(loc1, loc2);
		}
	}

	private void compareDataInSheet(Locator loc1, Locator loc2) {
		for (int j = 0; j < loc1.sheet.getPhysicalNumberOfRows(); j++) {
			if (loc2.sheet.getPhysicalNumberOfRows() <= j)
				return;

			loc1.row = loc1.sheet.getRow(j);
			loc2.row = loc2.sheet.getRow(j);
			if ((loc1.row == null) || (loc2.row == null)) {
				continue;
			}
			compareDataInRow(loc1, loc2);
		}
	}

	private void compareDataInRow(Locator loc1, Locator loc2) {
		for (int k = 0; k < loc1.row.getLastCellNum(); k++) {
			if (loc2.row.getPhysicalNumberOfCells() <= k)
				return;

			loc1.cell = loc1.row.getCell(k);
			loc2.cell = loc2.row.getCell(k);
			if ((loc1.cell == null) || (loc2.cell == null)) {
				continue;
			}
			compareDataInCell(loc1, loc2);
		}
	}

	private void compareDataInCell(Locator loc1, Locator loc2) {
		switch (loc1.cell.getCellType()) {
		case Cell.CELL_TYPE_BLANK:
		case Cell.CELL_TYPE_STRING:
		case Cell.CELL_TYPE_ERROR:
			isCellContentMatches(loc1, loc2);
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			isCellContentMatchesForBoolean(loc1, loc2);
			break;
		case Cell.CELL_TYPE_NUMERIC:
			if (DateUtil.isCellDateFormatted(loc1.cell)) {
				isCellContentMatchesForDate(loc1, loc2);
			} else {
				isCellContentMatchesForNumeric(loc1, loc2);
			}
			break;
		default:
			throw new IllegalStateException("Unexpected cell type: ");
		}
	}

	private void compareNumberOfColumnsInSheets(Locator loc1, Locator loc2) {
		for (int i = 0; i < loc1.workbook.getNumberOfSheets(); i++) {
			if (loc2.workbook.getNumberOfSheets() <= i)
				return;

			loc1.sheet = loc1.workbook.getSheetAt(i);
			loc2.sheet = loc2.workbook.getSheetAt(i);
			Iterator<Row> ri1 = loc1.sheet.rowIterator();
			Iterator<Row> ri2 = loc2.sheet.rowIterator();
			int num1 = (ri1.hasNext()) ? ri1.next().getPhysicalNumberOfCells() : 0;
			int num2 = (ri2.hasNext()) ? ri2.next().getPhysicalNumberOfCells() : 0;
			if (num1 != num2) {
				String str = loc1.sheet.getSheetName() + num1 + loc2.sheet.getSheetName() + num2;
				listOfDifferences.add(str);
			}
		}
	}

	private void compareNumberOfRowsInSheets(Locator loc1, Locator loc2) {
		for (int i = 0; i < loc1.workbook.getNumberOfSheets(); i++) {
			if (loc2.workbook.getNumberOfSheets() <= i)
				return;

			loc1.sheet = loc1.workbook.getSheetAt(i);
			loc2.sheet = loc2.workbook.getSheetAt(i);
			int num1 = loc1.sheet.getPhysicalNumberOfRows();
			int num2 = loc2.sheet.getPhysicalNumberOfRows();
			if (num1 != num2) {
				String str = loc1.sheet.getSheetName() + num1 + loc2.sheet.getSheetName() + +num2;
				listOfDifferences.add(str);
			}
		}

	}

	private void compareSheetData(Locator loc1, Locator loc2) {
		compareNumberOfRowsInSheets(loc1, loc2);
		compareNumberOfColumnsInSheets(loc1, loc2);
		compareDataInAllSheets(loc1, loc2);

	}

	private void isCellContentMatches(Locator loc1, Locator loc2) {
		String str1 = loc1.cell.getRichStringCellValue().getString();
		String str2 = loc2.cell.getRichStringCellValue().getString();
		if (!str1.equals(str2)) {
			addMessage(loc1, loc2, CELL_DATA_DOES_NOT_MATCH, str1, str2);
		}
	}

	private void isCellContentMatchesForBoolean(Locator loc1, Locator loc2) {
		boolean b1 = loc1.cell.getBooleanCellValue();
		boolean b2 = loc2.cell.getBooleanCellValue();
		if (b1 != b2) {
			addMessage(loc1, loc2, CELL_DATA_DOES_NOT_MATCH, Boolean.toString(b1), Boolean.toString(b2));
		}
	}

	private void isCellContentMatchesForDate(Locator loc1, Locator loc2) {
		Date date1 = loc1.cell.getDateCellValue();
		Date date2 = loc2.cell.getDateCellValue();
		if (!date1.equals(date2)) {
			addMessage(loc1, loc2, CELL_DATA_DOES_NOT_MATCH, date1.toString(), date2.toString());
		}
	}

	private void isCellContentMatchesForNumeric(Locator loc1, Locator loc2) {
		double num1 = loc1.cell.getNumericCellValue();
		double num2 = loc2.cell.getNumericCellValue();
		if (num1 != num2) {
			addMessage(loc1, loc2, CELL_DATA_DOES_NOT_MATCH, Double.toString(num1), Double.toString(num2));
		}
	}

	private void addMessage(Locator loc1, Locator loc2, String messageStart, String value1, String value2) {
		String str = messageStart + loc1.sheet.getSheetName() + new CellReference(loc1.cell).formatAsString() + value1
				+ loc2.sheet.getSheetName() + new CellReference(loc2.cell).formatAsString() + value2;
		listOfDifferences.add(str);
	}
}
