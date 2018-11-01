package XLCompare.com.santosh;


import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;



public class XLComparison {

	Properties r_XLProperties = null;
	//	Properties XL_Properties = null;
	static Xls_Reader xlRead = null;
	static String g_propertiesFileAddress = null;
	public static String g_DirectoryAddress = null;
	static String g_outputDataXLSFileName = null;

	public static Workbook workbook = null; 
	public Sheet g_sheet = null; 
	public Row g_row   =null;  
	String g_xlFileAddress =  null;
	String g_xlFileName = null;
	String g_resultFileName = null;
	static String g_Env1SheetName = null;
	static String g_Env2SheetName = null;
	static String g_resultSheetName = "Compared_Results";
	static String EVENT_REJECTION_INDICATOR_TEXT_IN_SHEET = "No Value found in 68... Event might be rejected.";
	
	int g_SourceIDColNo = 0;


	String [] primaryKeyColumn = null;
		
	XLComparison(){
		g_DirectoryAddress = System.getProperty("user.dir");
		g_propertiesFileAddress = g_DirectoryAddress+"/XLComparison.properties";
		g_outputDataXLSFileName = getProperty(g_propertiesFileAddress, "RESULT_FILE_NAME");


		g_xlFileAddress =  g_DirectoryAddress+"/FilesForComparison/";

		g_xlFileName = getProperty(g_propertiesFileAddress, "WORKBOOK_FILE_NAME");
		g_resultFileName = getProperty(g_propertiesFileAddress, "RESULT_FILE_NAME");
		g_Env1SheetName = getProperty(g_propertiesFileAddress, "ENV_1_SHEET_NAME");
		g_Env2SheetName = getProperty(g_propertiesFileAddress, "ENV_2_SHEET_NAME");
		
		xlRead = new Xls_Reader(g_xlFileAddress+g_xlFileName);
		
		primaryKeyColumn = getProperty(g_propertiesFileAddress, "PRIMARY_ATTRIBUTES").split(",");
		for(int i = 0; i<primaryKeyColumn.length; i++)
			System.out.println(primaryKeyColumn[i].toString());
		
		XLComparison.workbook = Xls_Reader.workbook;
		this.g_sheet = xlRead.sheet;
		this.g_row = xlRead.row;
	}

		
	/**
	 * Check whether the file exist in the given location or not
	 * @param fileAddress
	 * @return True if the passed parameter exists and is a file.
	 * ;  false if the passed parameter is either doesn't exist or a directory.
	 * @author Santosh Kumar
	 */
	public boolean isFileExist(String fileAddress) {

		File file = new File(fileAddress);
		if( file.exists() && file.isFile()) {
			//			System.out.println("File found in location "+ fileAddress);
			return true;
		}
		else if(file.isDirectory()) {
			System.out.println("The address "+ fileAddress+" is not a file but a directory");
		}
		else	
			System.out.println("File doesn't exist in location "+ fileAddress);
		return false;
	}

	public void findRangeAndFormatInAlternateColors(String sheetName, IndexedColors color1, IndexedColors color2 ) {
		int startRowNumOfRange = 0;
		String l_currentSourceID = null;
		String startSourceIDOfRange = null;
		Sheet currentSheet = workbook.getSheet(sheetName);
		int currentRowNum = 0;
		for(Row row:currentSheet) {
			currentRowNum = row.getRowNum();
			l_currentSourceID = xlRead.getCellValueInString(row.getCell(g_SourceIDColNo));
			if(currentRowNum>1) {
				if(startSourceIDOfRange.compareTo(l_currentSourceID)!=0) {
					xlRead.formatRange(currentSheet,startRowNumOfRange, currentRowNum-1,color1, color2);
					startRowNumOfRange=currentRowNum;
					startSourceIDOfRange = l_currentSourceID;
				}

				if(currentRowNum==currentSheet.getLastRowNum()) {
					xlRead.formatRange(currentSheet, startRowNumOfRange, currentRowNum,color1, color2 );
				}

			}else if(currentRowNum==1) {
				startRowNumOfRange = 1;
				startSourceIDOfRange = l_currentSourceID;
			}
		}
	}

	public String getProperty(String g_propertiesFileAddress, String propertyName) {
		r_XLProperties= new Properties();
		try {
			r_XLProperties.load(new FileInputStream(g_propertiesFileAddress));
		} catch (IOException e) {
			e.printStackTrace();
		}
		return r_XLProperties.getProperty(propertyName);
	}

	public File[] getAllFilesInDirectoryFromExtension(String pathOfDirectory, String extension) {
		//ToDo Use extension to fetch list of files.

		return  new File(pathOfDirectory).listFiles();
	}

	/*
	 * Open the file given in the file address, if found
	 * @param fileAddress
	 * @return true if file exist and successfully opened it; false otherwise
	 */
	public boolean openFile(String fileAddress) {
		if(isFileExist(fileAddress)) {
			File file = new File (fileAddress);
			//			file.setReadOnly();
			Desktop desktop = Desktop.getDesktop();
			try {
				desktop.open(file);
				return true;
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return false;
	}

	public void compareAndFillSecondEnvData(String secondEnvSheetName, String g_resultSheetName) {
		Sheet sheet2ndEnv = workbook.getSheet(secondEnvSheetName);
		Sheet resultSheet = workbook.getSheet(g_resultSheetName);
		for(Row rowIn2ndSheet:sheet2ndEnv) {
			if(rowIn2ndSheet.getRowNum()>0) {
				for(Row rowInResultSheet: resultSheet) {
					if(rowIn2ndSheet.getRowNum()>0) {
						if(xlRead.getCellValueInString(rowInResultSheet.getCell(0)).equals(EVENT_REJECTION_INDICATOR_TEXT_IN_SHEET)) {
							//if Matched: argument row number should be one lesser than no value row number
							if(xlRead.compareRowsUsingHeader(primaryKeyColumn,rowIn2ndSheet,resultSheet.getRow((rowInResultSheet.getRowNum())-1))) {
								xlRead.writeRowInSheet(rowIn2ndSheet, g_resultSheetName, rowInResultSheet.getRowNum());
								break;
							}

						}
					}
				}
			}
		}
	}

	public static void main(String[] args) {
		XLComparison o_XLComparison = new XLComparison();
		//		xlRead.compareSheetData(Xls_Reader.workbook.getSheetAt(0), Xls_Reader.workbook.getSheetAt(1),18);

//		System.out.println(g_DirectoryAddress);


		o_XLComparison.g_sheet= workbook.getSheet(g_Env1SheetName);

		//Create New Sheet and clone sheet 6 result to that sheet
		System.out.println("Cloning env 6 data to the result sheet...");
		try {
			xlRead.cloneSheetAndRename(workbook.getSheetIndex(workbook.getSheet(g_Env1SheetName)),g_resultSheetName);
		}catch(IllegalArgumentException e) {
			System.out.println("Exception occurs while cloning sheet");
			System.out.println("check file having given sheet "+g_Env1SheetName);
		}
		o_XLComparison.g_sheet=workbook.getSheet(g_resultSheetName);
		System.out.println("\t Cloning Successful\n\n");


		//Format Result sheet alternate colors for the files
		System.out.println("Formatting with alternate colors in result sheets based on Event file...");
		o_XLComparison.findRangeAndFormatInAlternateColors(g_resultSheetName,IndexedColors.LIGHT_CORNFLOWER_BLUE,IndexedColors.LIGHT_GREEN);
		System.out.println("\t Formatting Completed\n\n");


		//Insert alternate blank rows for 68 result place holding.
		System.out.println("Putting env 68 data in result sheet...");
		xlRead.insertAlternateBlankRows(g_resultSheetName, 1);


		//Compare the data based on the keys defined in primaryKeyColumn array 
		o_XLComparison.compareAndFillSecondEnvData(g_Env2SheetName, g_resultSheetName);
		System.out.println("\tEnv 68 data places successfully.\n\n");


		//Comparing Result 
		System.out.println("Comparing results now...");
		xlRead.compareAlternateRowData(workbook.getSheet(g_resultSheetName));
		System.out.println("\t Result compared successfully.\n\n");


		//Write and open file
		System.out.println("Writing the outcomes in new result file...");
		xlRead.writeFile(o_XLComparison.g_xlFileAddress+o_XLComparison.g_resultFileName);
		System.out.println("\t Written Successfully.\n\n");

		System.out.println("Opening the result file...");
		o_XLComparison.openFile(o_XLComparison.g_xlFileAddress+o_XLComparison.g_resultFileName);
		System.out.println("\tResult file opened successfully.");

	}

}
