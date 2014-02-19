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

public class Util {
	private static final int MY_MINIMUM_COLUMN_COUNT = 0;
//	Vector<EachSheet> eachSheet = new Vector<EachSheet>();
	Vector<Vector<EachSheet>> sheets = new Vector<Vector<EachSheet>>();
	void read(String fileName) throws InvalidFormatException, IOException {
		InputStream inp = new FileInputStream("input//" + fileName);
		Workbook wb = WorkbookFactory.create(inp);
		int sheetCount = wb.getNumberOfSheets();
//		初始化存放n个sheet的vector
		if (sheets.isEmpty()) {
			for (int i = 0; i < sheetCount; i++) {
				sheets.add(new Vector<EachSheet>());
			}
		}
		for (int i = 0; i < sheetCount; i++) {
			
			Sheet sheet = wb.getSheetAt(i);
			EachSheet thisEachSheet = new EachSheet();
			thisEachSheet.sheetname = sheet.getSheetName();
//			System.out.println(thisEachSheet.sheetname);
			int rowStart = sheet.getFirstRowNum();
			int rowEnd = sheet.getLastRowNum() + 1;
//			读取矩阵
			thisEachSheet.grid = new String[rowEnd][];
			for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
				Row r = sheet.getRow(rowNum);
				int lastColumn = Math.max(r.getLastCellNum(),
						MY_MINIMUM_COLUMN_COUNT);
				thisEachSheet.grid[rowNum] = new String[lastColumn];
				for (int cn = 0; cn < lastColumn; cn++) {
					Cell c = r.getCell(cn, Row.RETURN_BLANK_AS_NULL);
					if (c == null) {
						thisEachSheet.grid[rowNum][cn] = "";
					} else {
						thisEachSheet.grid[rowNum][cn] = c.toString();
					}
				}
			}
//			分析矩阵
//			得到列名
			thisEachSheet.headerName = new String[thisEachSheet.grid[0].length];
			for (int j = 0; j < thisEachSheet.grid[0].length; j++) {
				thisEachSheet.headerName[j] = thisEachSheet.grid[0][j];
			}
//			得到最后一列的值
			thisEachSheet.values.score = new String[thisEachSheet.grid.length];
			thisEachSheet.values.name = fileName.split(" ")[0];
			for (int j = 1; j < thisEachSheet.grid.length; j++) {
				thisEachSheet.values.score[j-1] = thisEachSheet.grid[j][thisEachSheet.grid[j].length - 1];
			}
			sheets.get(i).add(thisEachSheet);
		}
	}
	
	
	void output() throws IOException {
		Workbook wb = new XSSFWorkbook();
//		创建sheets
		for (int i = 0; i < sheets.size(); i++) {
			Sheet sheet = wb.createSheet(sheets.get(i).get(0).sheetname);
			System.out.print(sheets.get(i).get(0).sheetname);
//			创建行
			for (int j = 0; j < sheets.get(i).get(0).grid.length; j++) {
				Row row = sheet.createRow(j);
				for (int k = 0; k < sheets.get(i).get(0).grid[j].length - 1; k++) {
					row.createCell(k).setCellValue(sheets.get(i).get(0).grid[j][k].toString());
				}
				for (int k = 0; k < sheets.get(i).size(); k++) {
					if (j == 0) {
						row.createCell(sheets.get(i).get(0).grid[j].length - 1 + k).setCellValue(sheets.get(i).get(k).values.name.toString());
					}
					if (j != 0) {
						row.createCell(sheets.get(i).get(0).grid[j].length - 1 + k).setCellValue(sheets.get(i).get(k).values.score[j - 1].toString());	
					}
				}
			}

		}
		FileOutputStream fileOut = new FileOutputStream("output.xlsx");
		wb.write(fileOut);
		fileOut.close();
	}
}
