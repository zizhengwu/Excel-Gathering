package excelMaster;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.FlowLayout;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.File;
import java.io.IOException;
import java.io.PrintStream;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;

import javax.swing.JApplet;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextArea;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

@SuppressWarnings({ "serial" })
public class Main extends JApplet {
	Util util = new Util();

	public static void main(String[] args) throws InvalidFormatException,
			IOException {
		JFrame f = new JFrame("派点处理");
		Main mainFrame = new Main();
		f.getContentPane().add(mainFrame, BorderLayout.CENTER);
		f.pack();
		f.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		f.setLocationRelativeTo(null); // Center the frame
		f.setVisible(true);
		// TODO Auto-generated method stub
	}

	Main() throws UnsupportedEncodingException {
		// layout
		this.getContentPane().setLayout(new BorderLayout());
		this.getContentPane().add(util);
	}
}