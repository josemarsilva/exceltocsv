package org.exceltocsv.exceltocsv;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
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
	private Sheet worksheet;
	private HashMap<Integer, CellStyle> cellStyles = new HashMap<Integer, CellStyle>();


    /*
     * ExcelToCsv(cliArgsParser) - Constructor ...
     */
	public ExcelToCsv(CliArgsParser pCliArgsParser) {
		super();
		this.cliArgsParser = pCliArgsParser;
	}
	
	
}
