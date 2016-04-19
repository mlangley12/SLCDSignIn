import java.awt.EventQueue;
import  java.io.*;

import javax.swing.JFrame;
import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFCell;

import java.awt.GridLayout;

import javax.swing.JTextArea;

import java.awt.FlowLayout;

import javax.swing.JTextField;

import java.awt.Label;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;
import java.awt.Button;

import javax.swing.JMenuBar;
import javax.swing.JMenu;
import javax.swing.JMenuItem;

public class SignInGUI {

	private JFrame frame;
	private JOptionPane optionPane = new JOptionPane();
	private JTextField nameTextField;
	private JTextField schoolTextField;
	private JTextField emailTextField;
	private ArrayList<Attendees> attendees = new ArrayList<Attendees>();

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					SignInGUI window = new SignInGUI();
					window.frame.setVisible(true);			//make the gui
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public SignInGUI() {
		createWindow();
	}
	/**
	 * Initialize the contents of the frame.
	 */
	private void createWindow() {
		//create JFrame and contentPane
		frame = new JFrame();
		frame.setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
		frame.setBounds(100, 100, 450, 300);
		frame.getContentPane().setLayout(null);
		frame.addWindowListener(new WindowAdapter() {
			  public void windowClosing(WindowEvent e) {
			    int confirmed = JOptionPane.showConfirmDialog(null, 
			        "Are you sure you want to exit the program? All data will be lost.", "Exit Program Message Box",
			        JOptionPane.YES_NO_OPTION);

			    if (confirmed == JOptionPane.YES_OPTION) {
			      System.exit(0);
			    }
			  }
			});
		
		//create name textField
		nameTextField = new JTextField();
		nameTextField.setBounds(33, 135, 86, 20);
		frame.getContentPane().add(nameTextField);
		nameTextField.setColumns(10);
		
		//create school textField
		schoolTextField = new JTextField();
		schoolTextField.setBounds(178, 135, 86, 20);
		frame.getContentPane().add(schoolTextField);
		schoolTextField.setColumns(10);
		
		//create email textField
		emailTextField = new JTextField();
		emailTextField.setBounds(317, 135, 86, 20);
		frame.getContentPane().add(emailTextField);
		emailTextField.setColumns(10);
		
		//create name label
		Label label = new Label("Name(first, last)");
		label.setBounds(33, 107, 98, 22);
		frame.getContentPane().add(label);
		
		//create school label
		Label label_1 = new Label("School");
		label_1.setBounds(178, 107, 62, 22);
		frame.getContentPane().add(label_1);
		
		//create email label
		Label label_2 = new Label("Email");
		label_2.setBounds(317, 107, 62, 22);
		frame.getContentPane().add(label_2);
		
		//create submit button
		Button button = new Button("Submit");
		button.addActionListener(new ActionListener(){
				public void actionPerformed(ActionEvent e)
				{
					saveUser();
					nameTextField.setText(" ");
					schoolTextField.setText(" ");
					emailTextField.setText(" ");
					makeExcelDoc();
				}});
		button.setBounds(178, 197, 70, 22);
		frame.getContentPane().add(button);
		
		//create menu
		JMenuBar menuBar = new JMenuBar();
		frame.setJMenuBar(menuBar);
		JMenu mnFile = new JMenu("File");
		menuBar.add(mnFile);
		
		//Create option to create excel doc
		JMenuItem mntmCreateExcelDoc = new JMenuItem("Create Excel Doc");
		mnFile.add(mntmCreateExcelDoc);
		mnFile.addActionListener(new ActionListener(){
			public void actionPerformed(ActionEvent e)
			{
				makeExcelDoc();
			}});
	}
	//if a user hits enter it automatically saves them
	public void keyPressed( KeyEvent e ){
		if( e.getKeyCode() == KeyEvent.VK_ENTER ){
			saveUser();
			nameTextField.setText(" ");
			schoolTextField.setText(" ");
			emailTextField.setText(" ");
		}
	}
	
	//get the string the user entered in the name field
	public String getNameTextField(){
		return nameTextField.getText();
	}
	
	//get the string the user entered in the school field
	public String getSchoolTextField(){
		return schoolTextField.getText();
	}
	
	//get the string the user entered in the email field
	public String getEmailTextField(){
		return emailTextField.getText();
	}
	
	public void saveUser(){
		Attendees att = new Attendees(getNameTextField(), getSchoolTextField(), getEmailTextField());
		attendees.add(att);
	}
	//this print function is nice but we need a method that writes it to an excel file
	public void printList(){
		for(int i=0; i<attendees.size(); i++){
			System.out.println(attendees.get(i).getName() + " " + attendees.get(i).getSchool() + " " + attendees.get(i).getEmail());
		}
	}
	
	//public String getArrayObjectName(int n) {

	//}
	
	//create the excel doc with all objects in our arrayList attendees
	public void makeExcelDoc(){
		try {
            String filename = "E:/Eclipse Workspace/signInExcelDoc" ;
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("SignInExelDoc");  
            
            //create row titles
            HSSFRow rowhead = sheet.createRow((short)0);
            rowhead.createCell(0).setCellValue("Name");
            rowhead.createCell(1).setCellValue("Email");
            rowhead.createCell(2).setCellValue("School");
            
            //populate the rows with data in attendees arrayList
            HSSFRow row = sheet.createRow((short)1);
            for(int i=0; i < attendees.size(); i++){
            	Attendees att = attendees.get(i);
            	row.createCell(0).setCellValue(att.getName());
            	row.createCell(1).setCellValue(att.getEmail());
            	row.createCell(2).setCellValue(att.getSchool());
            }

            FileOutputStream fileOut = new FileOutputStream(filename);
            workbook.write(fileOut);
            fileOut.close();
            System.out.println("Your excel file has been generated!");

        } catch ( Exception ex ) {
            System.out.println(ex);
        }
	}
}
