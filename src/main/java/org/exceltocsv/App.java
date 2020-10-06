package org.exceltocsv;

import org.exceltocsv.cli.CliArgsParser;
import org.exceltocsv.exceltocsv.ExcelToCsv;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {

		// Parser Command Line Arguments ...
		CliArgsParser cliArgsParser = new CliArgsParser(args);
		
		// Create exceltocsv  instance ...
		ExcelToCsv excelToCsv = new ExcelToCsv(cliArgsParser);
		
        System.out.println( "Done!" );
    }
}
