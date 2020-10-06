package org.exceltocsv.cli;

import java.io.File;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.apache.commons.lang3.StringUtils;


/*
 * CliArgsParser class is responsible for extract, compile and check consistency
 * for command line arguments passed as parameters in command line
 * 
 * @author josemarsilva
 * @date   2020-10-06
 * 
 */
public class CliArgsParser {
	
	// Constants ...
	public final static String APP_NAME = new String("exceltocsv");
	public final static String APP_VERSION = new String("v.2020.10.06");
	public final static String APP_USAGE = new String(APP_NAME + " [<args-options-list>] - "+ APP_VERSION);

	// Constants defaults ...

	
	// Constants Error Messages ...
	public final static String MSG_ERROR_FILE_NOT_FOUND = new String("Erro: arquivo '%s' não existe !");
	public final static String MSG_ERROR_PATH_NOT_FOUND = new String("Erro: caminho da pasta '%s' não existe !");
	
	// Properties ...
    private String inputExcelFileOption = new String("");  // Nome do arquivo excel
    private String outputFolderPathCsvOption = new String("");  // Caminho o folder onde serao exportados os arquivos


    /*
     * CliArgsParser(args) - Constructor ...
     */
	public CliArgsParser( String[] args ) {

		// Options creating ...
		Options options = new Options();
		
		
		// Options configuring  ...
        Option helpOption = Option.builder("h") 
        		.longOpt("help") 
        		.required(false) 
        		.desc("shows usage help message. See more https://github.com/josemarsilva/exceltocsv") 
        		.build(); 
        Option inputExcelFileOption = Option.builder("i")
        		.longOpt("input-excel-file") 
        		.required(true) 
        		.desc("Nome do arquivo que contem Pasta de trabalho EXCEL (.xls ou .xlsx) que sera copiada para CSV. Ex: exemplo.xlsx")
        		.hasArg()
        		.build();
        Option outputFolderPathCsvOption = Option.builder("o")
                .longOpt("output-folder-path-csv") 
                .required(true) 
        		.desc("Nome do caminho da pasta onde deverao ser gerados os arquivos (.csv) correspondentes a cada sheet do workbook. Ex: C:\\TEMP")
        		.hasArg()
                .build(); 
        
        
		// Options adding configuration ...
        options.addOption(helpOption);
        options.addOption(inputExcelFileOption);
        options.addOption(outputFolderPathCsvOption);
        
        // CommandLineParser ...
        CommandLineParser parser = new DefaultParser();
		try {
			CommandLine cmdLine = parser.parse(options, args);
			
	        if (cmdLine.hasOption("help")) { 
	            HelpFormatter formatter = new HelpFormatter();
	            formatter.printHelp(APP_USAGE, options);
	            System.exit(0);
	        } else {
	        	
	        	// Set properties from Options ...
	        	this.setInputExcelFileOption( cmdLine.getOptionValue("input-excel-file", "") );	
	        	this.setOutputFolderPathCsvOption( cmdLine.getOptionValue("output-folder-path-csv", "") );	
	        	
	        	// Check arguments Options ...
	        	try {
	        		checkArgumentOptions();
	        	} catch (Exception e) {
	    			System.err.println(e.getMessage());
	    			System.exit(-1);
	        	}
	        	
	        }
			
		} catch (ParseException e) {
			System.err.println(e.getMessage());
            HelpFormatter formatter = new HelpFormatter(); 
            formatter.printHelp(APP_USAGE, options); 
			System.exit(-1);
		} 
        
	}

	//
	private void checkArgumentOptions() throws Exception {
		
		// Check input file arguments must exists ...
		if (!(new File(inputExcelFileOption).exists())) {
			throw new Exception(MSG_ERROR_FILE_NOT_FOUND.replaceFirst("%s", inputExcelFileOption));
		} else if (!(new File(outputFolderPathCsvOption).exists())) {
			throw new Exception(MSG_ERROR_PATH_NOT_FOUND.replaceFirst("%s", outputFolderPathCsvOption));
		}
		
	}
	
	// Getters and Setters ...

	public String getInputExcelFileOption() {
		return inputExcelFileOption;
	}


	public void setInputExcelFileOption(String inputExcelFileOption) {
		this.inputExcelFileOption = inputExcelFileOption;
	}

	public String getOutputFolderPathCsvOption() {
		return outputFolderPathCsvOption;
	}

	public void setOutputFolderPathCsvOption(String outputFolderPathCsvOption) {
		this.outputFolderPathCsvOption = outputFolderPathCsvOption;
	}


	//	outputExcelUtfEncoding

}
