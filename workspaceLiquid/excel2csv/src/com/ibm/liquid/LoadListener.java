/*
 * Copyright (C) 2014 IBM Corporation, All Rights Reserved.
 */
package com.ibm.liquid;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.JOptionPane;
import javax.swing.table.DefaultTableModel;

import org.apache.log4j.Logger;


/**
 * This is the Listener class that helps in processing the request coming from the Load button on the initial browse file window
 * 
 * @author Audas
 *
 */
public class LoadListener implements ActionListener {
	private static Logger LOGGER = Logger.getLogger(LoadListener.class);
	
	private XlsToCsvGui xlsToCsvGui;
	
	/**
	 * Constructor to accept XlsToCsvGui object
	 * @param xlsToCsvGui
	 */
	public LoadListener(XlsToCsvGui xlsToCsvGui) {
		this.xlsToCsvGui = xlsToCsvGui;
	}
	
	/* (non-Javadoc)
	 * @see java.awt.event.ActionListener#actionPerformed(java.awt.event.ActionEvent)
	 */
	public void actionPerformed(ActionEvent arg0) {
		LOGGER.info("actionPerformed IN");
		File selectedExcelFile = new File(xlsToCsvGui.getTextSelectExcel().getText());
		if (!(selectedExcelFile.exists() && selectedExcelFile.isFile())) {
			JOptionPane.showMessageDialog(xlsToCsvGui.getWindowMain(), "Please select a excel file before Load", "XlsToCsv", JOptionPane.PLAIN_MESSAGE);
			return;
		} else if(!selectedExcelFile.canRead()) {
			JOptionPane.showMessageDialog(xlsToCsvGui.getWindowMain(), "File is not readable", "XlsToCsv", JOptionPane.PLAIN_MESSAGE);
			return;
		}
		String[] header = {"Sheet name", "CSV file name", "Select"};
		LOGGER.debug("Excel file selected : " + xlsToCsvGui.getTextSelectExcel().getText());
		xlsToCsvGui.setXlsToCsvUtil(new XlsToCsvUtil(xlsToCsvGui.getTextSelectExcel().getText()));
		DefaultTableModel model = xlsToCsvGui.getXlsToCSVTableModel(xlsToCsvGui.getXlsToCsvUtil().getSheetMappings(), header);
		xlsToCsvGui.populateGridTable(model);
				
	}

}
