package excelMaster;

import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class Main {

	public static void main(String[] args) throws InvalidFormatException, IOException {
		// TODO Auto-generated method stub
		Util util = new Util();
		File dir = new File("input");
		for (File child : dir.listFiles()) {
			util.read(child.getName());
		}
		util.output();
	}

}
