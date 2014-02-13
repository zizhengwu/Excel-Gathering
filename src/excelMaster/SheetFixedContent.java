package excelMaster;

import java.util.Vector;

public class SheetFixedContent {
	private String sheetName;
	private int columnCount;
	private int rowCount;
	private Vector<String> columnHeader = new Vector<String>();
	private Vector<Vector<String>> Content = new Vector<Vector<String>>();
	public int getColumnCount() {
		return columnCount;
	}
	public void setColumnCount(int columnCount) {
		this.columnCount = columnCount;
	}
	public Vector<String> getColumnHeader() {
		return columnHeader;
	}
	public void setColumnHeader(Vector<String> columnHeader) {
		this.columnHeader = columnHeader;
	}
	public Vector<Vector<String>> getContent() {
		return Content;
	}
	public void setContent(Vector<Vector<String>> Content) {
		this.Content = Content;
	}
	public String getSheetName() {
		return sheetName;
	}
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}
	public int getRowCount() {
		return rowCount;
	}
	public void setRowCount(int rowCount) {
		this.rowCount = rowCount;
	}

}
