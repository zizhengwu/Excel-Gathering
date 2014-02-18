package excelMaster;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Vector;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
	private static final int MY_MINIMUM_COLUMN_COUNT = 0;

	Vector<SheetFixedContent> sheetFixedContent = new Vector<SheetFixedContent>();

	Vector<Vector<ArrayList<String>>> values = new Vector<Vector<ArrayList<String>>>();

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		Main main = new Main();
		main.init();
	}

	void init() throws Exception {

		this.readFormat("format.xlsx");
		File dir = new File("input");
		for (File child : dir.listFiles()) {
			this.readEvery(child.getName());
		}
		this.output("output.xlsx");
	}

	private void readEvery(String fileName) throws InvalidFormatException,
			IOException {
		
		InputStream inp = new FileInputStream("input//" + fileName);
		Workbook wb = WorkbookFactory.create(inp);
		for (int i = 0; i < sheetFixedContent.size(); i++) {

			Sheet sheet = wb.getSheetAt(i);
			ArrayList<String> thisValues = new ArrayList<String>();
			thisValues.add(fileName.split(" ")[0]);
			// Decide which rows to process
			int rowStart = sheet.getFirstRowNum();
			int rowEnd = sheet.getLastRowNum() + 1;

			for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
				Row r = sheet.getRow(rowNum);
				int lastColumn = Math.max(r.getLastCellNum(),
						MY_MINIMUM_COLUMN_COUNT);

				if (rowNum != rowStart) {
					Cell c = r
							.getCell(lastColumn - 1, Row.RETURN_BLANK_AS_NULL);
					if (c == null) {
						thisValues.add("");
					} else {
						thisValues.add(c.toString());
					}
				}
			}
			values.get(i).add(thisValues);
		}

	}

	private void output(String fileName) throws IOException {
		Workbook wb = new XSSFWorkbook();
		for (int i = 0; i < sheetFixedContent.size(); i++) {
			SheetFixedContent thisSheetFixedContent = sheetFixedContent.get(i);
			Sheet sheet = wb.createSheet(thisSheetFixedContent.getSheetName());

			// Create a row and put some cells in it. Rows are 0 based.
			for (int j = 0; j < thisSheetFixedContent.getRowCount(); j++) {
				Row row = sheet.createRow(j);
				if (j == 0) {
					for (int k = 0; k < thisSheetFixedContent.getColumnHeader().size(); k++) {
						row.createCell(k).setCellValue(
								thisSheetFixedContent.getColumnHeader().get(k));
					}
				} else {
					for (int k = 0; k < thisSheetFixedContent.getContent()
							.get(j).size(); k++) {
						row.createCell(k).setCellValue(
								thisSheetFixedContent.getContent().get(j)
										.get(k));
					}
				}
				for (int k = 0; k < values.get(i).size(); k++) {
					if (j == 0) {
						row.createCell(thisSheetFixedContent.getColumnCount() + k).setCellValue(values.get(i).get(k).get(j));
					}
					else {
						row.createCell(thisSheetFixedContent.getContent().get(j).size() + k).setCellValue(values.get(i).get(k).get(j));
					}
				}
				
			}
		}

		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream(fileName);
		wb.write(fileOut);
		fileOut.close();
	}

	void readFormat(String fileName) throws IOException, InvalidFormatException {
		InputStream inp = new FileInputStream("format.xlsx");
		Workbook wb = WorkbookFactory.create(inp);
		int sheetCount = wb.getNumberOfSheets();
		for (int i = 0; i < sheetCount; i++) {
			
			values.add(new Vector<ArrayList<String>>());
			SheetFixedContent thisSheetFixedContent = new SheetFixedContent();
			Sheet sheet = wb.getSheetAt(i);
			thisSheetFixedContent.setSheetName(sheet.getSheetName());
			System.out.println(sheet.getSheetName());
			// Decide which rows to process
			int rowStart = sheet.getFirstRowNum();
			int rowEnd = sheet.getLastRowNum() + 1;
			thisSheetFixedContent.setRowCount(rowEnd);

			for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
				Row r = sheet.getRow(rowNum);
				int lastColumn = Math.max(r.getLastCellNum(),
						MY_MINIMUM_COLUMN_COUNT);

				if (rowNum == rowStart) {
					thisSheetFixedContent.setColumnCount(lastColumn);
					for (int cn = 0; cn < lastColumn; cn++) {
						Cell c = r.getCell(cn, Row.RETURN_BLANK_AS_NULL);
						if (c != null) {
							thisSheetFixedContent.getColumnHeader().add(
									c.toString());
						}
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