package excelMaster;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.TextArea;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Vector;

import javax.swing.JButton;
import javax.swing.JPanel;
import javax.swing.JTextArea;
import javax.swing.JTextPane;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Util extends JPanel {
	TextArea log = new TextArea();
	RunButton run = new RunButton();
	private static final int MY_MINIMUM_COLUMN_COUNT = 0;
	// Vector<EachSheet> eachSheet = new Vector<EachSheet>();
	Vector<Vector<EachSheet>> sheets = new Vector<Vector<EachSheet>>();

	Util() {
		this.setLayout(new BorderLayout());
		this.setPreferredSize(new Dimension(600, 800));
		this.add(log, BorderLayout.CENTER);
		this.add(run, BorderLayout.SOUTH);
	}

	class RunButton extends JButton {
		public RunButton() {
			this.setText("Run");
			addMouseListener(new MouseAdapter() {
				public void mousePressed(MouseEvent e) {
					noClick();
					File dir = new File("input");
					for (File child : dir.listFiles()) {
						try {
							read(child.getName());
						} catch (InvalidFormatException e1) {
							// TODO Auto-generated catch block
							e1.printStackTrace();
						} catch (IOException e1) {
							// TODO Auto-generated catch block
							e1.printStackTrace();
						}
					}
					try {
						output();
					} catch (IOException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}
					check();
				}
			});
		}

		void noClick() {
			this.setEnabled(false);
			this.updateUI();
		}

		void read(String fileName) throws InvalidFormatException, IOException {
			log.append("正在读取: " + fileName + "\n");
			InputStream inp = new FileInputStream("input//" + fileName);
			Workbook wb = WorkbookFactory.create(inp);
			int sheetCount = wb.getNumberOfSheets();
			// 初始化存放n个sheet的vector
			if (sheets.isEmpty()) {
				for (int i = 0; i < sheetCount; i++) {
					sheets.add(new Vector<EachSheet>());
				}
			}
			for (int i = 0; i < sheetCount; i++) {

				Sheet sheet = wb.getSheetAt(i);
				EachSheet thisEachSheet = new EachSheet();
				thisEachSheet.sheetname = sheet.getSheetName();
				log.append(sheet.getSheetName() + "\t");
				// log.append(thisEachSheet.sheetname);
				int rowStart = sheet.getFirstRowNum();
				int rowEnd = sheet.getLastRowNum() + 1;
				// 读取矩阵
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
				// 分析矩阵
				// 得到列名
				thisEachSheet.headerName = new String[thisEachSheet.grid[0].length];
				for (int j = 0; j < thisEachSheet.grid[0].length; j++) {
					thisEachSheet.headerName[j] = thisEachSheet.grid[0][j];
				}
				// 得到最后一列的值
				thisEachSheet.values.score = new String[thisEachSheet.grid.length];
				thisEachSheet.values.name = fileName.split(" ")[0];
				for (int j = 1; j < thisEachSheet.grid.length; j++) {
					thisEachSheet.values.score[j - 1] = thisEachSheet.grid[j][thisEachSheet.grid[j].length - 1];
				}
				sheets.get(i).add(thisEachSheet);
				log.append("\n");
			}
		}

		void output() throws IOException {
			log.append("开始输出!" + "\n");
			Workbook wb = new XSSFWorkbook();
			// 创建sheets
			for (int i = 0; i < sheets.size(); i++) {
				Sheet sheet = wb.createSheet(sheets.get(i).get(0).sheetname);
				log.append("正在输出Sheet: " + sheets.get(i).get(0).sheetname
						+ "\n");
				// 创建行
				for (int j = 0; j < sheets.get(i).get(0).grid.length; j++) {
					Row row = sheet.createRow(j);
					for (int k = 0; k < sheets.get(i).get(0).grid[j].length - 1; k++) {
						row.createCell(k).setCellValue(
								sheets.get(i).get(0).grid[j][k].toString());
					}
					for (int k = 0; k < sheets.get(i).size(); k++) {
						if (j == 0) {
							row.createCell(
									sheets.get(i).get(0).grid[j].length - 1 + k)
									.setCellValue(
											sheets.get(i).get(k).values.name
													.toString());
						}
						if (j != 0) {
							try {
								row.createCell(
										sheets.get(i).get(0).grid[j].length - 1 + k)
										.setCellValue(
												sheets.get(i).get(k).values.score[j - 1]
														.toString());
							} catch (NullPointerException e) {
								log.append(sheets.get(i).get(k).values.name + "\t");
							}
							
						}
					}
				}

			}
			FileOutputStream fileOut = new FileOutputStream("output.xlsx");
			wb.write(fileOut);
			fileOut.close();
			log.setText("");;
			log.append("看起来成功了！\n");
		}

		void check() {
			for (int i = 0; i < sheets.size(); i++) {
				for (int j = 0; j < sheets.get(i).size(); j++) {
					float shouldLowerThan100 = 0;
					for (int k = 0; k < sheets.get(i).get(j).values.score.length; k++) {
						if (sheets.get(i).get(j).values.score[k] != "") {
							try {
								float thisScore = Float.parseFloat(sheets
										.get(i).get(j).values.score[k]);
								if (thisScore > 30) {
									log.append("!大于30："
											+ sheets.get(i).get(j).values.name
											+ " "
											+ sheets.get(i).get(j).sheetname
											+ "\n");
								}
								shouldLowerThan100 += thisScore;
							} catch (Exception e) {
								// TODO: handle exception
							}
						}
					}
					if (shouldLowerThan100 > 100) {
						log.append("!列之和大于100："
								+ sheets.get(i).get(j).values.name + " "
								+ sheets.get(i).get(j).sheetname + "\n");
					}
				}
			}
		}

	}
}
