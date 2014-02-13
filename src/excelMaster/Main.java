package excelMaster;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Vector;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
	private static final int MY_MINIMUM_COLUMN_COUNT = 0;

	Vector<SheetFixedContent> sheetFixedContent = new Vector<SheetFixedContent>();

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		Main main = new Main();
		main.init();
	}

	void init() throws Exception {
		this.readFormat("CJ.xls");
		this.output("output.xls");
	}

	private void output(String fileName) throws IOException {
		Workbook wb = new HSSFWorkbook();
		for (int i = 0; i < sheetFixedContent.size(); i++) {
			SheetFixedContent thisSheetFixedContent = sheetFixedContent.get(i);
			Sheet sheet = wb.createSheet(thisSheetFixedContent.getSheetName());

			// Create a row and put some cells in it. Rows are 0 based.
			for (int j = 0; j < thisSheetFixedContent.getRowCount(); j++) {
				Row row = sheet.createRow(j);
				for (int k = 0; k < thisSheetFixedContent.getContent().get(j).size(); k++) {
					row.createCell(k).setCellValue(
							thisSheetFixedContent.getContent().get(j).get(k));
					System.out.println(thisSheetFixedContent.getContent().get(j).get(k));
				}
			}
		}

		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream(fileName);
		wb.write(fileOut);
		fileOut.close();
	}

	void readFormat(String fileName) throws IOException, InvalidFormatException {
		InputStream inp = new FileInputStream(fileName);
		Workbook wb = WorkbookFactory.create(inp);
		int sheetCount = wb.getNumberOfSheets();
		System.out.println("sheetCount = " + sheetCount);
		for (int i = 0; i < sheetCount; i++) {
			System.out.println("Sheet " + i);
			SheetFixedContent thisSheetFixedContent = new SheetFixedContent();
			Sheet sheet = wb.getSheetAt(i);
			thisSheetFixedContent.setSheetName(sheet.getSheetName());
			// Decide which rows to process
			int rowStart = sheet.getFirstRowNum();
			int rowEnd = sheet.getLastRowNum() + 1;
			thisSheetFixedContent.setRowCount(rowEnd);

			for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
				System.out.println("rowNum = " + rowNum);
				Row r = sheet.getRow(rowNum);
				System.out.println(r.getLastCellNum());
				int lastColumn = Math.max(r.getLastCellNum(),
						MY_MINIMUM_COLUMN_COUNT);

				if (rowNum == rowStart) {
					thisSheetFixedContent.setColumnCount(lastColumn);
					for (int cn = 0; cn < lastColumn; cn++) {
						Cell c = r.getCell(cn, Row.RETURN_BLANK_AS_NULL);

						thisSheetFixedContent.getColumnHeader().add(
								c.toString());

					}
					for (int j = 0; j < thisSheetFixedContent.getRowCount(); j++) {
						thisSheetFixedContent.getContent().add(
								new Vector<String>());
					}
				} else {
					for (int cn = 0; cn < lastColumn; cn++) {
						Cell c = r.getCell(cn, Row.RETURN_BLANK_AS_NULL);
						if (c == null) {
							thisSheetFixedContent.getContent().get(rowNum)
									.add("");
						} else {
							thisSheetFixedContent.getContent().get(rowNum)
									.add(c.toString());
						}
					}
				}

			}
			sheetFixedContent.add(thisSheetFixedContent);
		}
	}

}
