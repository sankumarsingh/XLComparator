package XLCompare.com.santosh;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.piccolo.io.FileFormatException;



public class Xls_Reader { 
	//	public static String filename = "/src/com/santosh/res/XLFiles/Comparison_Batch1.xlsx"; 
	public  String fileAddress; 
	public  FileInputStream fis = null; 
	public  FileOutputStream fileOut =null; 
	public static Workbook workbook = null; 
	public Sheet sheet = null; 
	public Row row   =null; 
	public Sheet testSuiteSheet = null;
	public Sheet currentTestCaseSheet = null;
	public boolean bColor = false;
	public short colorIndex = 10;
	public CellStyle errorCellStyle = null;


	/**
	 * get Cell value as String 
	 * Cell type includes Boolean, Numeric, String, Formula, or Blank
	 * @param cell
	 * @return
	 * @author kumsanto
	 */
	public String getCellValueInString(Cell cell) {

		try {
			String cellValue="";

			CellType cellType = cell.getCellTypeEnum();

			switch (cellType) {
			case BOOLEAN:
				cellValue = cell.getBooleanCellValue() +"";
				break;
			case NUMERIC:
				cellValue = cell.getNumericCellValue() + "";
				break;
			case STRING:
				cellValue = cell.getStringCellValue();
				break;
			case FORMULA:
				cellValue = cell.getCellFormula();
				break;
			case BLANK:
				break;
			default:
				break;
			}
			return cellValue;
		}catch(NullPointerException e) {
			System.out.println("Cell is null");
			return "";
		}
	}


	/**
		/**
	 * Load the excel file from the fileAddress. Accepts xlsx and xls files only else will show error
	 * In case file type is .xlsx, it will initialize will XSSFWorkbook
	 * In case file type is .xls, it will initialize will HSSFWorkbook
	 * In case the file type is neither xls nor xlsx it will throw FileFormatException
	 * @param fileAddress
	 * @author kumsanto
	 */
	public void loadExcelFile(String fileAddress) { 
		this.fileAddress=fileAddress; 
		if(isFileExist(fileAddress)) {
			//			setReadOnly(fileAddress);
			try { 
				fis = new FileInputStream(fileAddress); 
				if(getExcelExtension(fileAddress).equalsIgnoreCase(".xlsx")) {
					ZipSecureFile.setMinInflateRatio(-1.0d);
					workbook = new XSSFWorkbook(fis); 
				}
				else if(getExcelExtension(fileAddress).equalsIgnoreCase(".xls")) {
					workbook = new HSSFWorkbook(fis); 
				}else if(getExcelExtension(fileAddress).equalsIgnoreCase("")) {
					throw new FileFormatException("FileFormatIsNotCorrect");
				}
				sheet = workbook.getSheetAt(0); 
				row = sheet.getRow(0);
				row.getCell(0);
				fis.close(); 
			} catch (FileFormatException e) { 
				System.out.println("Not an excel file type " + e.getMessage());
				e.printStackTrace(); 
			}catch (Exception e) { 
				e.printStackTrace(); 
			} 
		}
		else
			System.out.println("File Address is not correct.");
	} 


	/**
	 * Constructor used to initialize the xls_reader
	 * @param fileAddress
	 * @author Santosh Kumar
	 */
	public Xls_Reader(String fileAddress) { 
		System.out.println(fileAddress);
		loadExcelFile(fileAddress);
		initiateErrorCellStyle();
		//other for init
	} 
	public void setFocusOnSheet(String SheetName) {
		workbook.getSheet(SheetName).getRow(0).getCell(0).setAsActiveCell();

	}


	public void shiftRowsByNNumber(String sheetName, int fromRow, int shiftByNNumber) {
		Sheet sheet = workbook.getSheet(sheetName);
		sheet.shiftRows(fromRow,sheet.getLastRowNum() , shiftByNNumber);
	}

	/**
	 * Return column number that contains the header text in header row.
	 * @param sheet
	 * @param headerRowNum
	 * @param header
	 * @return column number of Header
	 */
	public int getHeaderColumnFromText(Sheet sheet, int headerRowNum, String header) {
		Row headerRow = sheet.getRow(headerRowNum);
		int num =-1;
		for (int cellNum=0;cellNum<headerRow.getLastCellNum();cellNum++) {
			if(headerRow.getCell(cellNum).getStringCellValue().equalsIgnoreCase(header.trim())) {
				num = cellNum;
				break;
			}
		}
		return num;
	}


	/**
	 * Check the file extension is .xlsx or xls or other file type.
	 * @param fileAddress
	 * @return ".xls" if filetype is .xls;
	 * 			".xlsx" if filetype is .xlsx;
	 * 			blank if file type is neither .xlsx, nor .xls
	 * @author Santosh Kumar
	 */
	public String getExcelExtension(String fileAddress) {
		if(fileAddress.toUpperCase().endsWith(".XLSX")) {
			return ".xlsx";
		}
		else if(fileAddress.toUpperCase().endsWith(".XLS")) {
			return ".xls";
		}

		return "";
	}

	public void findColumnAndWriteDataInCell(int RowNum, String Header, String data) {
		Row header = sheet.getRow(RowNum);
		for(Cell cell: header) {
			if(Header.equalsIgnoreCase(cell.getStringCellValue())){
				int CellNumber = cell.getColumnIndex();
				sheet.getRow(RowNum).getCell(CellNumber).setCellValue(data);
			}
		}
	}

	/**

	 * Check whether the file exist in the given location or not
	 * @param fileAddress
	 * @return True if the passed parameter exists and is a file.
	 * ; Return false if the passed parameter is either doesn't exist or a directory.
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


	/**
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
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		return false;
	}

	/**
	 * Write the excel file in the given file address
	 * @param fileAddress
	 * @return true if successfully write the file; false otherwise
	 * @author kumsanto
	 * 
	 */
	public boolean writeFile(String fileAddress) {
		try {
			File file = new File (fileAddress);
			//			file.setWritable(true);
			//			if (file.canWrite()) {
			FileOutputStream fout = new FileOutputStream(file);
			workbook.write(fout);
			fout.close();
			//			}else
			//				System.out.println("***Unable to Write in the file... Please close all of its open instances");
			//			file.setWritable(false);

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}


		return false;
	}

	public void compareRow_ColumnByColumn(Row row1, Row row2) {
		System.out.println(row.getPhysicalNumberOfCells());
		for (Cell cell1:row1) {
			Cell cell2 = row2.getCell(cell1.getColumnIndex());
			if(getCellValueInString(cell1).compareTo(getCellValueInString(cell2))!=0 && cell1.getColumnIndex()>27) {
				highlightCell(cell1);
				highlightCell(cell2);
			}
		}
	}

	public void compareRow_ColumnByColumn(Row row1, Row row2,int fromColumnNumber) {
		for (Cell cell1:row1) {
			Cell cell2 = row2.getCell(cell1.getColumnIndex());
			if(getCellValueInString(cell1).compareTo(getCellValueInString(cell2))!=0) {
				if(cell1.getColumnIndex()>fromColumnNumber) {
					//					highlightCell(cell1);
					CellStyle cellstyle1 = cell1.getCellStyle();
					cellstyle1.setFillForegroundColor(IndexedColors.RED.getIndex());
					cellstyle1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					cell1.setCellStyle(cellstyle1);
					cell2.setCellStyle(cellstyle1);

				}
			}
		}
	}

	public void compareRow_ColumnByColumnInSheet(Row row1, Row row2,int fromColumnNumber,String resultSheet) {
		//		System.out.println("Comparing Row no "+row1.getRowNum());

		for (Cell cell1:row1) {
			Cell cell2 = row2.getCell(cell1.getColumnIndex());
			if(getCellValueInString(cell1).compareTo(getCellValueInString(cell2))!=0) {
				if(cell1.getColumnIndex()>fromColumnNumber) {
					highlightCell(cell1);
					highlightCell(cell2);
				}
			}
		}
	}

	public void colorRow(Row row, IndexedColors bgColor, IndexedColors fontColor) {
		CellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(bgColor.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		//		style.setFillPattern(CellStyle.SOLID_FOREGROUND); 
		Font font = workbook.createFont();
		font.setColor(fontColor.getIndex());
		style.setFont(font);
		for(Cell cell:row)
			cell.setCellStyle(style);


	}

	public void createNewSheetByDeletingExistingSheet(String sheetName) {

		if(workbook.getSheetIndex(sheetName)>=0) 
			workbook.removeSheetAt(workbook.getSheetIndex(sheetName));
		workbook.createSheet(sheetName);
	}

	public void checkAndCreateNewSheet(String sheetName) {
		if(workbook.getSheetIndex(sheetName)<0) {
			workbook.createSheet(sheetName);
		}else
			System.out.println(sheetName+" sheet already present");
	}


	public void initiateErrorCellStyle() {
		errorCellStyle = workbook.createCellStyle();
		errorCellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
		errorCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		Font font = workbook.createFont();
		font.setColor(IndexedColors.BLUE.getIndex());
		errorCellStyle.setFont(font);

	}

	public void highlightCell(Cell cell) {
		//		System.out.println("Cell ("+cell.getRowIndex()+", "+cell.getColumnIndex()+") is highlighted.");
		CellStyle style = workbook.createCellStyle();
		style.cloneStyleFrom(cell.getCellStyle());
		style.setFillForegroundColor(IndexedColors.RED.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cell.setCellStyle(style);
	}

	public CellStyle createCellStyle(IndexedColors indexColor) {

		CellStyle style = workbook.createCellStyle();
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFillForegroundColor(indexColor.getIndex());
		return style;
	}

	public CellStyle updateCellStyle(IndexedColors indexColor, CellStyle oldCellStyle) {

		CellStyle style = workbook.createCellStyle();
		style.cloneStyleFrom(oldCellStyle);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFillForegroundColor(indexColor.getIndex());
		return style;
	}

	public void writeRowInSheet(Row sourceRow, String destinationSheetName, int rowNumberInDestination) {
		workbook.getSheet(destinationSheetName).createRow(rowNumberInDestination);

		for(Cell sourceCell:sourceRow) {

			CellStyle destCellStyle = workbook.createCellStyle();
			destCellStyle.cloneStyleFrom(sourceCell.getCellStyle());
			//			System.out.println("\t"+sourceCell.getColumnIndex());
			Cell destCell = workbook.getSheet(destinationSheetName).
					getRow(rowNumberInDestination).
					createCell(sourceCell.getColumnIndex());

			//			destCell.setCellValue(getCellValueInString(sourceCell));
			destCell.setCellType(sourceCell.getCellTypeEnum());
			destCell.setCellStyle(destCellStyle);

			switch(sourceCell.getCellTypeEnum()){
			case  NUMERIC :
				destCell.setCellType(CellType.NUMERIC);
				destCell.setCellValue(sourceCell.getNumericCellValue());
				break;
			case STRING:
				destCell.setCellType(CellType.STRING);
				destCell.setCellValue(sourceCell.getStringCellValue());
				break;
			case BLANK:
				destCell.setCellType(CellType.BLANK);
				destCell.setCellValue("");
				break;
			default:
				break;

			}
			try{
				if(DateUtil.isCellDateFormatted(sourceCell)) {
					Date date = new Date();
					date = sourceCell.getDateCellValue();
					destCell.setCellValue(date);
				}
			}catch(IllegalStateException e) {
				e.getMessage();
			}
		}
	}

	public void writeBlankRowInSheet(String destinationSheetName, int rowNumberInDestination) {
		workbook.getSheet(destinationSheetName).createRow(rowNumberInDestination);
		for(Cell cell:workbook.getSheet(destinationSheetName).getRow(0)) {
			workbook.getSheet(destinationSheetName).getRow(rowNumberInDestination).createCell(cell.getColumnIndex()).setCellValue("");
		}
	}

	public int getColumnIndex(Sheet sheet, String columnHeader, int headerRowNum) {
		Row row = sheet.getRow(headerRowNum);
		int columnIndex = 0;
		for(Cell cell:row) {
			if(getCellValueInString(cell).equalsIgnoreCase(columnHeader)) {
				columnIndex =	 cell.getColumnIndex();
				break;
			}
		}
		return columnIndex;
	}

	public void compareSheetData(Sheet sheet1, Sheet sheet2, int fromColumnNumber) {

		for (Row row1:sheet1) {
			Row row2=sheet2.getRow(row1.getRowNum());
			compareRow_ColumnByColumn(row1, row2,fromColumnNumber);


		}
	}

	public void cloneSheetAndRename(int sheetNumber, String resultSheetName) throws IllegalArgumentException {
		//		TODO If sheet already exist will throw IllegalArgumentException
		workbook.setSheetName(workbook.getSheetIndex(workbook.cloneSheet(sheetNumber)), resultSheetName);

	}

	public void insertAlternateBlankRows(String sheetName, int fromRow) {
		//		System.out.println("Inserting Alternate Blank Rows");
		sheet=workbook.getSheet(sheetName);
		int rowNum = sheet.getLastRowNum();
		for(;rowNum>0;rowNum--) {
			if(rowNum!=sheet.getLastRowNum()) {
				sheet.shiftRows(rowNum+1, sheet.getLastRowNum()+1, 1);
				Row row = sheet.createRow(rowNum+1);
				row.createCell(0).setCellValue(XLComparison.EVENT_REJECTION_INDICATOR_TEXT_IN_SHEET);

			}else if(rowNum==sheet.getLastRowNum()) {
				sheet.createRow(rowNum+1).createCell(0).setCellValue(XLComparison.EVENT_REJECTION_INDICATOR_TEXT_IN_SHEET);
			}

		}
	}

	public boolean compareRowsUsingHeader(String headers[], Row row1, Row row2) {
		int currentColumn;
		for(String header:headers) {
			currentColumn = getHeaderColumnFromText(row1.getSheet(), 0, header);

			//check if any value is not same;
			if(getCellValueInString(row1.getCell(currentColumn)).compareTo(getCellValueInString(row2.getCell(currentColumn)))!=0){
				//				return false;
				break;
			}
		}

//		System.out.println("Row1 "+row1.getRowNum()+" and Row2 "+row2.getRowNum()+ "are same");
		return true;
	}

	public void compareAlternateRowData(Sheet sheet) {
		for(int rowNum =1; rowNum<sheet.getLastRowNum();rowNum = rowNum+2) {
			if(!getCellValueInString(sheet.getRow(rowNum+1).getCell(0)).contains(XLComparison.EVENT_REJECTION_INDICATOR_TEXT_IN_SHEET)) {
				compareRow_ColumnByColumn(sheet.getRow(rowNum),sheet.getRow(rowNum+1),18);
			}
		}
	}

	public void formatRange(Sheet sheet, int startRowNum, int endRowNum,IndexedColors color1, IndexedColors color2 ) {
		IndexedColors currentColor = null;
		if(bColor==false) {
			currentColor = color1;
			bColor=true;
		}else {
			currentColor = color2;
			bColor=false;
		}
		//		System.out.println("Formatting ranges from "+startRowNum+" to "+endRowNum);
		for (int rowNum = startRowNum; rowNum<= endRowNum; rowNum++) {
			row = sheet.getRow(rowNum);
			try {
				for(Cell cell:row) {
					//					CellStyle style = cell.getCellStyle();

					cell.setCellStyle(updateCellStyle(currentColor,cell.getCellStyle()));

					//					cell.setCellStyle(style);

				}
			}catch(NullPointerException e) {
				e.getMessage();
			}
		}

	}

	public void formatRange(Sheet sheet, int startRowNum, int endRowNum ) {
		IndexedColors currentColor = null;
		if(bColor==false) {
			currentColor = IndexedColors.LIGHT_CORNFLOWER_BLUE;
			bColor=true;
		}else {
			currentColor = IndexedColors.LIGHT_GREEN;
			bColor=false;
		}
		for (int rowNum = startRowNum; rowNum<= endRowNum; rowNum++) {
			row = sheet.getRow(rowNum);
			try {
				for(Cell cell:row) {
					cell.setCellStyle(updateCellStyle(currentColor,cell.getCellStyle()));
				}
			}catch(NullPointerException e) {
				e.getMessage();
			}
		}

	}
}	
