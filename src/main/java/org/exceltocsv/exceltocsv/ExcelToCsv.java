package org.exceltocsv.exceltocsv;

import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.exceltocsv.cli.CliArgsParser;


public class ExcelToCsv {
	
	// Constants Error Messages ...
	public final static String MSG_ERROR_EXCEL_EXCEPTION = new String("Erro: exception '%s' durante operação '%s' com parametros '%s'!");

	
	// Properties ...
	private CliArgsParser cliArgsParser;
	
	private Workbook workbook;


    /*
     * ExcelToCsv(cliArgsParser) - Constructor ...
     */
	public ExcelToCsv(CliArgsParser pCliArgsParser) {
		super();
		this.cliArgsParser = pCliArgsParser;
	}


	/*
	 * executeExportToCsv() - Execute export to (.csv)
	 */
	public void executeExportToCsv() {

		// Open Workbook ...
		System.out.println("Opening workbook '%s' ...".replaceFirst("%s",this.cliArgsParser.getInputExcelFileOption()));
		try {
			this.workbook = WorkbookFactory.create(new FileInputStream(this.cliArgsParser.getInputExcelFileOption().replace("\\", "//")));
		} catch (Exception e) {
			// e.printStackTrace();
			System.err.println(MSG_ERROR_EXCEL_EXCEPTION.replaceFirst("s", e.getMessage()).replaceFirst("%s", "WorkbookFactory.create()").replaceFirst("%s", this.cliArgsParser.getInputExcelFileOption().replace("\\", "//") ));
			System.exit(-1);
		}
		
		// Loop iterate WorkSheets ...
		Iterator<Sheet> sheetIterator = workbook.iterator();
		while (sheetIterator.hasNext()) {
		    Sheet sheet = sheetIterator.next();
		    
		    // Output (.csv) filename ...
		    String outputCsvFilename = new String("");
		    outputCsvFilename = new String( sheet.getSheetName().replace(" ", "").concat(".csv") );
		    outputCsvFilename = new String( this.cliArgsParser.getOutputFolderPathCsvOption().concat("\\").concat("\\").concat(outputCsvFilename) );
			System.out.println("Exporting worksheet '%s' to '%s' ...".replaceFirst("%s",sheet.getSheetName()).replaceFirst("%s",outputCsvFilename));

			// FileWriter and BufferWriter ... 
			FileWriter fileWriter = null;
			try {
				fileWriter = new FileWriter(outputCsvFilename);
			} catch (IOException e) {
				// e.printStackTrace();
				System.err.println(MSG_ERROR_EXCEL_EXCEPTION.replaceFirst("s", e.getMessage()).replaceFirst("%s", "new FileWriter()") );
				System.exit(-1);
			}
			BufferedWriter bufferedWriter = new BufferedWriter(fileWriter);
			try {

				// Loop iterate Rows  ...
				for (Row row : sheet) {
					String strRow = new String("");
					// Loop iterate Cells  ...
				    for (Cell cell : row) {
						String strCell = new String("");
						// Cell is not empty
						if (cell != null) {
							switch (cell.getCellType()) {
							case Cell.CELL_TYPE_BLANK:
								strCell = new String("");
								break;
							case Cell.CELL_TYPE_BOOLEAN:
								strCell = ( cell.getBooleanCellValue() ) ? new String("True") : new String("False");
								break;
							case Cell.CELL_TYPE_ERROR:
								strCell = new String("#ERROR");
								break;
							case Cell.CELL_TYPE_FORMULA:
								strCell = new String(cell.getStringCellValue());
								break;
							case Cell.CELL_TYPE_NUMERIC:
								if (HSSFDateUtil.isCellDateFormatted(cell)) {
									Date cellDate = cell.getDateCellValue();
									DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");  
									strCell =  dateFormat.format(cellDate);
									if (strCell.length()>10) {
										if (strCell.substring(11,strCell.length()).contentEquals("00:00:00")) {
											strCell = new String(strCell.substring(0,10));
										}
									}
								} else {
									strCell = doubleToStringFmt(cell.getNumericCellValue());
								}
								break;
							case Cell.CELL_TYPE_STRING:
								strCell = new String(cell.getStringCellValue()).replace("\n", "");
								break;
							default:
								break;
							}
						}
						if (strRow.isEmpty()) {
							strRow = new String(strCell);
						} else {
							strRow = new String(strRow.concat(";").concat(strCell));
						}
				    }
				    
				    // Write bufferedWriter ...
				    bufferedWriter.write(strRow.concat("\n"));
				    
				}

			} catch (IOException e) {
				// e.printStackTrace();
				System.err.println(MSG_ERROR_EXCEL_EXCEPTION.replaceFirst("s", e.getMessage()).replaceFirst("%s", "bufferedWriter.write()") );
				System.exit(-1);
			}
			try {
				bufferedWriter.close();
			} catch (IOException e) {
				// e.printStackTrace();
				System.err.println(MSG_ERROR_EXCEL_EXCEPTION.replaceFirst("s", e.getMessage()).replaceFirst("%s", "bufferedWriter.close()") );
				System.exit(-1);
			}
		}

		// Close Workbook ...
		try {
			workbook.close();
		} catch (Exception e) {
			// e.printStackTrace();
			System.err.println(MSG_ERROR_EXCEL_EXCEPTION.replaceFirst("s", e.getMessage()).replaceFirst("%s", "workbook.close()") );
			System.exit(-1);
		}

	}

	
	/*
	 * doubleToStringFmt() :: String - Convert and format a Double value to String value using brute force
	 */
	public String doubleToStringFmt(double d) {
		String strIntegerPortion = new String("");
		String strDecimalPortion = new String("");
		if (d<Integer.MAX_VALUE) {
	        int integerValue = (int) d;
	        double decimalValue = d - (double) integerValue;
	        strIntegerPortion = Integer.toString(integerValue);
	        if (decimalValue != (double) 0) {
	        	strDecimalPortion = Double.toString(decimalValue);
	        	strDecimalPortion = strDecimalPortion.substring(1,strDecimalPortion.length()-1);
	        	if (strDecimalPortion.length()>3 && strDecimalPortion.substring(3,strDecimalPortion.length()).replace("0", "").contentEquals("") ) {
		        	strDecimalPortion = strDecimalPortion.substring(0,3+1);
	        	}
	        }
		} else {
			strDecimalPortion = new String(Double.toString(d));
			if (strDecimalPortion.contains("E")) {
				int times10 = 0;
				times10 = Integer.valueOf( strDecimalPortion.substring(strDecimalPortion.indexOf("E")+1,strDecimalPortion.length()) );
				strDecimalPortion = new String(strDecimalPortion.substring(0,strDecimalPortion.indexOf("E")).replace(".", ""));
				for(int i=strDecimalPortion.length()-1;i<times10;i++) {
					strDecimalPortion = new String(strDecimalPortion.concat("0"));
				}
			}
		}
		return strIntegerPortion.concat(strDecimalPortion);
	}

	
}
