/*
 * Copyright (C) 2014 IBM Corporation, All Rights Reserved.
 */
package com.ibm.liquid;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import javax.swing.JFileChooser;
import javax.swing.filechooser.FileFilter;
import org.apache.log4j.Logger;

/**
 * 
 * Browse button listener to open File Chooser dialogue
 * 
 * @author Audas
 *
 */
public class BrowseListener implements ActionListener {
	private static Logger LOGGER = Logger.getLogger(BrowseListener.class);
	
	private XlsToCsvGui xlsToCsvGui;
	
	/**
	 * Constructor to accept XlsToCsvGui object
	 * 
	 * @param xlsToCsv
	 */
	public BrowseListener(XlsToCsvGui xlsToCsvGui) {
		this.xlsToCsvGui = xlsToCsvGui;
	}

	
	/* (non-Javadoc)
	 * @see java.awt.event.ActionListener#actionPerformed(java.awt.event.ActionEvent)
	 */
	public void actionPerformed(ActionEvent arg0) {
		LOGGER.info("actionPerformed IN");
		LOGGER.debug(arg0.getActionCommand());

		if (arg0.getActionCommand().equalsIgnoreCase("BrowseOutputPath")) {
			xlsToCsvGui.getOutputPathChooser().setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
			xlsToCsvGui.getOutputPathChooser().setMultiSelectionEnabled(false);
			int returnVal = xlsToCsvGui.getOutputPathChooser().showOpenDialog(xlsToCsvGui.getWindowMain());
			if (returnVal == JFileChooser.APPROVE_OPTION) {
				xlsToCsvGui.getTextOutputPath().setText(xlsToCsvGui.getOutputPathChooser().getSelectedFile().getAbsolutePath());
			}
		} else if (arg0.getActionCommand().equalsIgnoreCase("BrowseExcelFile")) {	
			xlsToCsvGui.getInputFileChooser().setFileSelectionMode(JFileChooser.FILES_ONLY);
			xlsToCsvGui.getInputFileChooser().setMultiSelectionEnabled(false);
			xlsToCsvGui.getInputFileChooser().setFileFilter(new ExtensionFileFilter("xls", new String[] { "xls" }));
			int returnVal = xlsToCsvGui.getInputFileChooser().showOpenDialog(xlsToCsvGui.getWindowMain());
			if (returnVal == JFileChooser.APPROVE_OPTION) {
				xlsToCsvGui.getTextSelectExcel().setText(xlsToCsvGui.getInputFileChooser().getSelectedFile().getAbsolutePath());
			}

		} else {
			LOGGER.debug("Invalid command");
		}
		
		return;
	}


}

/**
 * <p>This class is for filter of file chooser</p>
 * 
 * @author Audas
 *
 */
class ExtensionFileFilter extends FileFilter {
	String description;

	String extensions[];

	/**
	 * @param description
	 * @param extension
	 */
	public ExtensionFileFilter(String description, String extension) {
		this(description, new String[] { extension });
	}

	/**
	 * @param description
	 * @param extensions
	 */
	public ExtensionFileFilter(String description, String extensions[]) {
		if (description == null) {
			this.description = extensions[0];
		} else {
			this.description = description;
		}
		this.extensions = (String[]) extensions.clone();
		toLower(this.extensions);
	}

	/**
	 * @param array
	 */
	private void toLower(String array[]) {
		for (int i = 0, n = array.length; i < n; i++) {
			array[i] = array[i].toLowerCase();
		}
	}

	/* (non-Javadoc)
	 * @see javax.swing.filechooser.FileFilter#getDescription()
	 */
	public String getDescription() {
		return description;
	}

	/* (non-Javadoc)
	 * @see javax.swing.filechooser.FileFilter#accept(java.io.File)
	 */
	@Override
	public boolean accept(File file) {
		if (file.isDirectory()) {
			return true;
		} else {
			String path = file.getAbsolutePath().toLowerCase();
			for (int i = 0, n = extensions.length; i < n; i++) {
				String extension = extensions[i];
				if ((path.endsWith(extension) && (path.charAt(path.length()
						- extension.length() - 1)) == '.')) {
					return true;
				}
			}
		}
		return false;
	}
}