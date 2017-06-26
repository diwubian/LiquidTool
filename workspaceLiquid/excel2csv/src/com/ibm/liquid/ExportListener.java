/*
 * Copyright (C) 2014 IBM Corporation, All Rights Reserved.
 */
package com.ibm.liquid;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.HashMap;
import java.util.Map;

import javax.swing.JOptionPane;
import javax.swing.table.TableModel;
import org.apache.log4j.Logger;


/**
 * This is the Listener class that helps in processing the request coming from the Export button on the application GUI
 * 
 * @author Audas
 *
 */
public class ExportListener implements ActionListener {
	private static Logger LOGGER = Logger.getLogger(ExportListener.class);
	
	private XlsToCsvGui xlsToCsvGui;
	
	/**
	 * Constructor to accept XlsToCsvGui object
	 * @param xlsToCsvGui
	 */
	public ExportListener(XlsToCsvGui xlsToCsvGui) {
		this.xlsToCsvGui = xlsToCsvGui;
	}
	
	/* (non-Javadoc)
	 * @see java.awt.event.ActionListener#actionPerformed(java.awt.event.ActionEvent)
	 */
	public void actionPerformed(ActionEvent arg0) {
		LOGGER.info("actionPerformed IN");
		
		if (xlsToCsvGui.getXlsToCsvUtil() == null) {
			LOGGER.error("Please Load before Export");
			JOptionPane.showMessageDialog(xlsToCsvGui.getWindowMain(), "Please Load before Export", "xlsToCsv", JOptionPane.ERROR_MESSAGE);
			return;
		}
		
		boolean error = false;
		
		TableModel table = xlsToCsvGui.getGridTable().getModel();
		Map<String, String> mappings = new HashMap<String, String>();
		for (int rowIndex=0; rowIndex<table.getRowCount(); rowIndex++) {
			if ((Boolean)table.getValueAt(rowIndex, 2))
				mappings.put(table.getValueAt(rowIndex, 0).toString(), table.getValueAt(rowIndex, 1).toString());
		}
		if (mappings.size() == 0) {
			LOGGER.debug("No sheet selected to export");
			JOptionPane.showMessageDialog(xlsToCsvGui.getWindowMain(), "Please select at least one sheet to export", "XlsToCsv", JOptionPane.PLAIN_MESSAGE);
			return;
		}
		if (xlsToCsvGui.getTextOutputPath().getText() == null || xlsToCsvGui.getTextOutputPath().getText().trim().equals("")) {
			LOGGER.debug("Output path not selected; Csv will be generated at current path");
			error = xlsToCsvGui.getXlsToCsvUtil().export(mappings, "./");
		} else {
			LOGGER.debug("Output path selected " + xlsToCsvGui.getTextOutputPath().getText());
			error = xlsToCsvGui.getXlsToCsvUtil().export(mappings, xlsToCsvGui.getTextOutputPath().getText());
		}
		String message;
		if (error) 
			message = "Export command done with error(s); Please check the log file";
		else
			message = "Export command done successfully.";
		JOptionPane.showMessageDialog(xlsToCsvGui.getWindowMain(), message, "XlsToCsv", JOptionPane.PLAIN_MESSAGE);
				
	}

}
