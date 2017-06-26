/*
 * Copyright (C) 2014 IBM Corporation, All Rights Reserved.
 */
package com.ibm.liquid;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.Toolkit;
import java.util.HashMap;
import java.util.Map;
import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.border.BevelBorder;
import javax.swing.border.EtchedBorder;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableModel;
import org.apache.log4j.Logger;

/**
 * This is the main java swing GUI component class that constructs the initial browse file window and with the help of 
 * other helper classes it opens the choose file dialog and finally the sheet to CSV mapping 
 * 
 * @author Audas
 *
 */
public class XlsToCsvGui {
	private static Logger LOGGER = Logger.getLogger(XlsToCsvGui.class);
	
	/**
	 * XlsToCsvUtil
	 */
	XlsToCsvUtil xlsToCsvUtil;
	
	/**
	 * Default constants to format the column & row of the Grid/JTable
	 */
	private static final int ROWHEIGHT = 25;

	/**
	 * Sheet name -> CSV file mapping
	 */
	public static final Map<String, String> SHEET_CSV_MAPPING = new HashMap<String, String>(); {
		SHEET_CSV_MAPPING.put("HW MT",		"hwmt.csv");
		SHEET_CSV_MAPPING.put("HW Mod",		"hwmod.csv");           
		SHEET_CSV_MAPPING.put("HW FC",		"hwfc.csv");
		SHEET_CSV_MAPPING.put("FA MT",		"famt.csv");
		SHEET_CSV_MAPPING.put("FA FC",		"fafc.csv");
		SHEET_CSV_MAPPING.put("Drive FC",		"ddfc.csv");
		SHEET_CSV_MAPPING.put("Functions",		"func.csv");
		SHEET_CSV_MAPPING.put("Functions - Scope",	"funclsopt.csv");
		SHEET_CSV_MAPPING.put("Functions - Sort",	"funcsort.csv");
		SHEET_CSV_MAPPING.put("Functions - Override","funcovr.csv");
		SHEET_CSV_MAPPING.put("Roles",		"roles.csv");
		SHEET_CSV_MAPPING.put("User Functions",	"functions.csv");
		SHEET_CSV_MAPPING.put("Roles - Functions",	"role_functions.csv");
		SHEET_CSV_MAPPING.put("Roles - Reports",	"role_reports.csv");
		SHEET_CSV_MAPPING.put("Log Retention",	"audited_functions.csv");		
	}
	
	/**
	 * Window /JFrame of the application
	 */
	private JFrame windowMain;
	
	/**
	 * Main panel of the Frame. Other Panels/Components will be added on this panel
	 */
	private JPanel mainPanel;
	
	/**
	 * JTable to show the Table data to select Work Item
	 */
	private JTable gridTable;	

	/**
	 * Label to show information
	 */
	private JLabel labelSelectExcel;
	
	/**
	 * Text field to display selected file with path
	 */
	private JTextField textSelectExcel;	
	
	/**
	 * Button to browse the file system to select Excel files
	 */
	private JButton buttonInputFile;
	
	/**
	 * File chooser dialogue
	 */
	private JFileChooser inputFileChooser;
		
	/**
	 * Button to browse the file system to select output path
	 */
	private JButton buttonOutputpath;
	
	/**
	 * CSV Output path
	 */
	private JLabel labelOutputPath;
	
	/**
	 * Text field to get output path
	 */
	private JTextField textOutputPath;
	
	/**
	 * File chooser dialogue
	 */
	private JFileChooser outputPathChooser;
	
	/**
	 * Load button to display Sheet information n JTable
	 */
	private JButton buttonLoad;
		
	/**
	 * Export to Csv button
	 */
	private JButton buttonExport;	

	/**
	 * @return JFrame
	 */
	public JFrame getWindowMain() {
		return windowMain; 
	}
	
	/**
	 * @return JPanel
	 */
	public JPanel getMainPanel() {
		return mainPanel;
	}	


	/**
	 * @return the gridTable
	 */
	public JTable getGridTable() {
		return gridTable;
	}		
	

	/**
	 * @return the xlsToCsvUtil
	 */
	public XlsToCsvUtil getXlsToCsvUtil() {
		return xlsToCsvUtil;
	}

	/**
	 * 
	 * @param xlsToCsvUtil
	 */
	public void setXlsToCsvUtil(XlsToCsvUtil xlsToCsvUtil) {
		this.xlsToCsvUtil = xlsToCsvUtil;
	}
	
	/**
	 * @return the labelOutputPath
	 */
	public JLabel getLabelOutputPath() {
		return labelOutputPath;
	}

	/**
	 * @return the labelSelectExcel
	 */
	public JLabel getLabelSelectExcel() {
		return labelSelectExcel;
	}

	/**
	 * @return the textOutputPath
	 */
	public JTextField getTextOutputPath() {
		return textOutputPath;
	}

	/**
	 * Constructs first screen of CSV selection
	 */
	public void displayUI() {
		LOGGER.info("displayUI IN");
		inputFileChooser = new JFileChooser();
		outputPathChooser = new JFileChooser();
		
		windowMain = new JFrame();
		windowMain.setTitle("Export to CSV");
		mainPanel =  new JPanel();
		mainPanel.setLayout(new GridBagLayout());
		
		GridBagConstraints gridBagConstraints = new GridBagConstraints();		
		labelSelectExcel = new JLabel("Excel file");	 
		labelSelectExcel.setBorder(BorderFactory.createEtchedBorder(EtchedBorder.LOWERED));
		gridBagConstraints.fill = GridBagConstraints.HORIZONTAL;
		gridBagConstraints.insets = new Insets(20, 10, 0, 10);
		gridBagConstraints.weightx = 0;
		gridBagConstraints.gridx = 0;
		gridBagConstraints.gridy = 0;
		gridBagConstraints.gridwidth = 1;
		labelSelectExcel.setPreferredSize(new Dimension(100, 25));		
		mainPanel.add(labelSelectExcel, gridBagConstraints);
		  
		textSelectExcel = new JTextField();		
		textSelectExcel.setForeground(Color.red);
		textSelectExcel.setBackground(Color.white);
		textSelectExcel.setBorder(BorderFactory.createEtchedBorder(EtchedBorder.LOWERED));
		textSelectExcel.setEditable(false);
		gridBagConstraints.fill = GridBagConstraints.HORIZONTAL;
		gridBagConstraints.weightx = 0;
		gridBagConstraints.gridx = 1;
		gridBagConstraints.gridy = 0;
		gridBagConstraints.gridwidth = 1;		
		textSelectExcel.setPreferredSize(new Dimension(0, 25));		
		mainPanel.add(textSelectExcel, gridBagConstraints);

		buttonInputFile = new JButton("Browse");
		gridBagConstraints.fill = GridBagConstraints.HORIZONTAL;
		gridBagConstraints.weightx = 0;
		gridBagConstraints.gridx = 2;
		gridBagConstraints.gridy = 0;
		buttonInputFile.setActionCommand("BrowseExcelFile");
		buttonInputFile.addActionListener(new BrowseListener(this));
		mainPanel.add(buttonInputFile, gridBagConstraints);
		  		
		buttonLoad = new JButton("Load");		
		gridBagConstraints.fill = GridBagConstraints.HORIZONTAL;
		gridBagConstraints.weightx = 0;
		gridBagConstraints.gridx = 3;
		gridBagConstraints.gridy = 0;		
		buttonLoad.addActionListener(new LoadListener(this));		
		mainPanel.add(buttonLoad, gridBagConstraints);
		
		labelOutputPath = new JLabel("Output path");	 
		labelOutputPath.setBorder(BorderFactory.createEtchedBorder(EtchedBorder.LOWERED));
		gridBagConstraints.fill = GridBagConstraints.HORIZONTAL;
		gridBagConstraints.insets = new Insets(6, 10, 6, 10);
		gridBagConstraints.weightx = 0;
		gridBagConstraints.gridx = 0;
		gridBagConstraints.gridy = 1;
		gridBagConstraints.gridwidth = 1;
		labelOutputPath.setPreferredSize(new Dimension(100, 25));
		mainPanel.add(labelOutputPath, gridBagConstraints);
		  
		textOutputPath = new JTextField();
		textOutputPath.setForeground(Color.red);
		textOutputPath.setBackground(Color.white);
		textOutputPath.setBorder(BorderFactory.createEtchedBorder(EtchedBorder.LOWERED));
		textOutputPath.setEditable(false);
		gridBagConstraints.fill = GridBagConstraints.HORIZONTAL;
		gridBagConstraints.weightx = 0.7;
		gridBagConstraints.gridx = 1;
		gridBagConstraints.gridy = 1;
		gridBagConstraints.gridwidth = 1;		
		textOutputPath.setPreferredSize(new Dimension(0, 25));		
		mainPanel.add(textOutputPath, gridBagConstraints);

		buttonOutputpath = new JButton("Browse");
		gridBagConstraints.fill = GridBagConstraints.HORIZONTAL;
		gridBagConstraints.weightx = 0;
		gridBagConstraints.gridx = 2;
		gridBagConstraints.gridy = 1;
		buttonOutputpath.setActionCommand("BrowseOutputPath");
		buttonOutputpath.addActionListener(new BrowseListener(this));
		mainPanel.add(buttonOutputpath, gridBagConstraints);
		
		buttonExport = new JButton("Export");
		gridBagConstraints.fill = GridBagConstraints.HORIZONTAL;
		gridBagConstraints.weightx = 0;
		gridBagConstraints.gridx = 3;
		gridBagConstraints.gridy = 1;
		gridBagConstraints.gridwidth = 1;
		buttonExport.addActionListener(new ExportListener(this));		
		mainPanel.add(buttonExport, gridBagConstraints);		
		
		windowMain.getContentPane().add(mainPanel);
		
		gridBagConstraints.fill = GridBagConstraints.HORIZONTAL;
		gridBagConstraints.insets = new Insets(6, 10, 6, 10);
		gridBagConstraints.anchor = GridBagConstraints.NORTHWEST;
		gridBagConstraints.weightx = 1;
		gridBagConstraints.weighty = 1;
		gridBagConstraints.gridx = 0;
		gridBagConstraints.gridy = 2;
		gridBagConstraints.gridwidth = 3;
		gridBagConstraints.ipady=180;
		gridTable = new JTable();
		mainPanel.add(new JScrollPane(gridTable), gridBagConstraints);		
		
		showWindow(windowMain);
	}
	
		
	/**
	 * @return the textSelectExcel
	 */
	public JTextField getTextSelectExcel() {
		return textSelectExcel;
	}

	/**
	 * @return the inputFileChooser
	 */
	public JFileChooser getInputFileChooser() {
		return inputFileChooser;
	}
	
	/**
	 * @return the OutputPathChooser
	 */
	public JFileChooser getOutputPathChooser() {
		return outputPathChooser;
	}
	
	/**
	 * This method populates Grid Table
	 * 
	 * @param model
	 * @param result
	 */
	public void populateGridTable(TableModel model) {
		LOGGER.info("populateGridTable IN");
		
		gridTable.setModel(model);
		gridTable.setRowHeight(ROWHEIGHT);
		gridTable.getTableHeader().setPreferredSize(new Dimension(gridTable.getWidth(), ROWHEIGHT));

		gridTable.getTableHeader().setBackground(Color.GRAY);
		gridTable.getTableHeader().setForeground(Color.YELLOW);
		gridTable.getTableHeader().setFont(new Font(gridTable.getTableHeader().getFont().getFamily(), Font.BOLD, gridTable.getTableHeader().getFont().getSize()));
		gridTable.getTableHeader().setBorder(new BevelBorder(BevelBorder.RAISED));
		gridTable.setAutoResizeMode(JTable.AUTO_RESIZE_ALL_COLUMNS);
		gridTable.setAutoResizeMode(JTable.WIDTH);		

	}

	
	/**
	 * Setting JFrame as per Window area & window type
	 * 
	 * @param frame
	 */
	private void showWindow(JFrame frame) {
		LOGGER.info("showWindow IN");
		
		try{
	        Toolkit tk = Toolkit.getDefaultToolkit();
	        Dimension screenSize = tk.getScreenSize();
	        int screenHeight = screenSize.height - 50;
	        int screenWidth = screenSize.width;
	    	frame.setTitle("Export sheets in one Excel to CSV files");
	    	frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
	    	frame.setVisible(true);
	        frame.setSize((int)(screenWidth/1.5), (int)(screenHeight/2));
	        frame.setLocation(0, 0); 
	        frame.setResizable(false);
	        frame.setVisible(true);
	        
		} catch(Throwable t) {
			LOGGER.error(t);
		}
	}

	/**
	 * getXlsToCSVTableModel
	 * 
	 * @param data
	 * @param header
	 * @return XlsToCSVTableModel
	 */
	public XlsToCSVTableModel getXlsToCSVTableModel(Object[][] data, String[] header) {
		return new XlsToCSVTableModel(data, header);
		
	}
		
	/**
	 * The JTable Model class to manage cell read only
	 * 
	 * @author Audas
	 *
	 */
	private class XlsToCSVTableModel extends DefaultTableModel {

		private static final long serialVersionUID = 5422344318449878720L;
		
		/**
		 * @param data
		 * @param header
		 */
		public XlsToCSVTableModel(Object[][] data, String[] header) {
			super(data, header);
		}
		
		@Override
		public boolean isCellEditable(int row, int column) {
			if (column == 0 )
				return false;
			else
				return true;
		}
		
		@Override
		public Class<?> getColumnClass(int columnIndex) {
			 switch (columnIndex) {
             case 0:
                 return String.class;
             case 1:
                 return String.class;
              default:
                 return Boolean.class;
         }
		}
		
		
	}
	
	
	/**
	 * The Main method for starting the application.
	 * 
	 * @param args
	 */
	public static void main(String[] args) {
		LOGGER.info("Starting GUI");
		XlsToCsvGui xlsToCSV = null;
		try {
			xlsToCSV = new XlsToCsvGui();
			xlsToCSV.displayUI();
		} catch (Exception e) {
			LOGGER.error(e);
			JOptionPane.showMessageDialog(xlsToCSV.getWindowMain(), e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
		}
	}

}

