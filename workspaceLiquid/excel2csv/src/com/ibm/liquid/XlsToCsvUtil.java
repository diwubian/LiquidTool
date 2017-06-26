/*
 * Copyright (C) 2014 IBM Corporation, All Rights Reserved.
 */
package com.ibm.liquid;

import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

/**
 * @author Audas
 *
 */

public class XlsToCsvUtil {
	private static Logger LOGGER = Logger.getLogger(XlsToCsvUtil.class);
	
	/**
	 * CSV_DELIMETER
	 */
	private static final String CSV_DELIMETER = ",";
	
	private String excelFile;
	private HSSFWorkbook workbook;
	private List<HSSFSheet> sheets;
	
	/**
	 * Constructor
	 * 
	 * @param excelFile
	 */
	public XlsToCsvUtil(String excelFile) {
		this.excelFile = excelFile;
	}
	
	/**
	 * This method returns List of Sheet name
	 * 
	 * @return
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	private List<String> getSheetNames() throws FileNotFoundException, IOException {
		LOGGER.info("getSheetNames IN");
		workbook = new HSSFWorkbook(new FileInputStream(excelFile));
		sheets =  new ArrayList<HSSFSheet>();
		List<String> sheetNameList = new ArrayList<String>();
		for (int i=0; i<workbook.getNumberOfSheets(); i++) {
			sheets.add(workbook.getSheetAt(i));
			sheetNameList.add(workbook.getSheetAt(i).getSheetName());
		}

		return sheetNameList;
	}
	
	/**
	 * This method returns Sheet name to CSV file mapping
	 * 
	 * @return
	 */
	public Object[][] getSheetMappings() {
		LOGGER.info("getSheetMappings IN");
		Object[][] retmap = new Object[0][3];
		try {
			List<String> list = getSheetNames();
			List<Object[]> maplist = new ArrayList<Object[]>();
			
			for (Iterator<String> iterator = list.iterator(); iterator.hasNext();) {
				String sheetname = iterator.next();
				if (XlsToCsvGui.SHEET_CSV_MAPPING.containsKey(sheetname))
					maplist.add(new Object[]{sheetname, XlsToCsvGui.SHEET_CSV_MAPPING.get(sheetname), Boolean.TRUE});
				
			}
			
			retmap = new Object[maplist.size()][3];
			for (int i = 0; i < maplist.size(); i++) {
				retmap[i] = maplist.get(i);
			}
			
			
		} catch (FileNotFoundException e) {
			LOGGER.error(e);
		} catch (IOException e) {
			LOGGER.error(e);
		}
				
		return retmap;
		
	}

	/**
	 * This method exports the sheets to corresponding CSV file
	 * 
	 * @param mappings
	 * @param outputPath
	 * @return boolean
	 */
	public boolean export(Map<String, String> mappings, String outputPath) {
		LOGGER.info("export IN");
		
		boolean error = false;
		
		BufferedWriter writer = null;
		String csvFilePath;
		HSSFSheet sheet = null;
		HSSFRow row = null;
		HSSFCell cell = null;
		String cellValue = "";
		int maxColCount = 0;
		
		for (Iterator<HSSFSheet> iterator = sheets.iterator(); iterator.hasNext();) {
			sheet = iterator.next();
			
			try {
				if (mappings.containsKey(sheet.getSheetName()))
					if (outputPath == null || outputPath.trim().equals("")) {
						LOGGER.debug("Output path not selected");
						csvFilePath = mappings.get(sheet.getSheetName());
					} else {
						csvFilePath = outputPath + "/" + mappings.get(sheet.getSheetName());
					}				
				else {
					LOGGER.debug(sheet.getSheetName() + " not included in mapping; skipping");
					continue;
				}
				LOGGER.debug(sheet.getSheetName() + " is going to exported in " + csvFilePath);
				writer = new BufferedWriter(new FileWriter(csvFilePath));

				maxColCount = getMaxCellNo(sheet);
				for (int rowno=0; rowno<=sheet.getLastRowNum(); rowno++) {
	
					row = sheet.getRow(rowno);
					if (row != null) {
						for (int colno=0; colno<maxColCount; colno++) {
							
							cell = row.getCell(colno);
							cellValue = "";
							if (cell != null) {
								switch (cell.getCellType()) {
								case Cell.CELL_TYPE_NUMERIC:
									cellValue = trimTrailingZeros(cell.toString());
									break;

								default:
									cellValue = escapeIt(cell.toString());
									break;
								}
							}
							if (colno > 0) {
								writer.write(CSV_DELIMETER + cellValue);
							} else {
								writer.write(cellValue);
							}
						}
					}
									
					writer.newLine();
				}
				LOGGER.debug(sheet.getSheetName() + " exported successfully in " + csvFilePath);
				
				
			} catch (IOException e) {
				LOGGER.error(e);
				error = true;
			} finally {
				if (writer != null) {
					try {
						writer.close();
					} catch (IOException e) {
						LOGGER.error(e);
					}
				}
			}
			
		}
		
		return error;
		
	}
	
	/**
	 * returns maximum column count of a sheet
	 * 
	 * @param sheet
	 * @return
	 */
	private int getMaxCellNo(HSSFSheet sheet) {
		HSSFRow row = null;
		int maxColumnCount = 0;
		for (int rowno=0; rowno<=sheet.getLastRowNum(); rowno++) {
			row = sheet.getRow(rowno);
			if (row != null) {
				if (row.getLastCellNum() > maxColumnCount)
					maxColumnCount = row.getLastCellNum();
				}
		}
		return maxColumnCount;
	}
	
	/**
	 * Trim trailing zeroes
	 * 
	 * @param number
	 * @return
	 */
	private String trimTrailingZeros(String number) {
	    if(!number.contains(".")) {
	        return number;
	    }

	    return number.replaceAll("\\.?0*$", "");
	}
	
	/**
	 * Escape special character 
	 * 
	 * @param value
	 * @return
	 */
	private String escapeIt(String value) {
		if (value.contains(CSV_DELIMETER) || value.contains("\"") || value.contains("\n")) {
			value = "\"" + value + "\""; 
		}
		return value;
	}
	

}
