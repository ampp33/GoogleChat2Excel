package org.malibu.convert.gchat;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GoogleChatToExcelConverter {
	
	private static final int TIME_COLUMN_INDEX = 0;
	private static final int NAME_COLUMN_INDEX = 1;
	private static final int TEXT_COLUMN_INDEX = 2;
	
	private static final String MINUTES_DIVIDER_REGEX = "^[\\s]*[\\d]+ minute[s]*[\\s]*$";
	private static final String TIME_REGEX = "^[\\s]*([\\d]{1,2}:\\d\\d [PA]M).*$";
	private static final String NAME_REGEX = "^[\\s]*(?:[\\d]{1,2}:\\d\\d [PA]M)?[\\s]*([a-zA-Z][0-9]:|[mM][eE]:).+$";
	private static final String TEXT_REGEX = "^[\\s]*(?:[\\d]{1,2}:\\d\\d [PA]M)?[\\s]*(?:[a-zA-Z][0-9]:|[mM][eE]:)?(.+)$";
	
	private static CellStyle MINUTES_DIVIDER_CELL_STYLE = null;
	private static CellStyle TIME_CELL_STYLE = null;
	private static CellStyle NAME_CELL_STYLE = null;
	private static CellStyle TEXT_CELL_STYLE = null;
	
	private ChatElementType lastElementType = null;
	private String lastName = null;
	private int currentRow = 1;
	
	private Workbook workbook = null;
	
	public static void main(String[] args) {
		GoogleChatToExcelConverter converter = new GoogleChatToExcelConverter();
		// get list of all files in the same dir as the jar file
		File jarFileLocation = new File(getJarLocationPath());
//		System.out.println(jarFileLocation.getAbsolutePath());
//		if(true) {
//			return;
//		}
		File[] jarFileDirFileList = jarFileLocation.listFiles();
		for (File file : jarFileDirFileList) {
			if(file == null) continue;
			String excelFilePath = null;
			try {
				// attempt to convert the current file
				excelFilePath = converter.convert(file.getAbsolutePath());
			} catch (IOException e) {
				System.err.println("Error occurred converting file: " + file.getAbsolutePath());
				e.printStackTrace();
				continue;
			}
			if(excelFilePath == null) {
				System.err.println("Didn't attempt to convert file: " + file.getAbsolutePath());
			} else {
				System.out.println("Successfully converted '" + file.getAbsolutePath() + "' to '" + excelFilePath + "'");
			}
		}
	}
	
	public static String getJarLocationPath() {
		// get directory .jar file is running from (using substring() to remove leading slash)
		String workingDir = GoogleChatToExcelConverter.class.getProtectionDomain().getCodeSource().getLocation().getPath();
		File file = new File(workingDir);
		workingDir = file.getAbsolutePath();
		if(workingDir.startsWith("\\")) {
			workingDir = workingDir.substring(1);
		}
		if(workingDir.endsWith(".")) {
			workingDir = workingDir.substring(0, workingDir.length() - 2);
		}
		return workingDir;
	}
	
	/**
	 * Converts the supplied Google chat text file to a spreadsheet
	 * 
	 * @param gChatFilePath
	 * @return
	 * @throws IOException
	 */
	public String convert(String gChatFilePath) throws IOException {
		// reset starting variables
		currentRow = 0;
		lastElementType = null;
		
		File gChatFile = new File(gChatFilePath);
		String outputExcelFilePath = null;
		if(gChatFile.exists() && gChatFile.isFile() && hasSpecificExtension(gChatFilePath, "txt")) {
			initWorkbook();
			BufferedReader stream = new BufferedReader(new FileReader(gChatFile));
			String line = null;
			while((line = stream.readLine()) != null) {
				processLine(line);
			}
			stream.close();
			outputExcelFilePath = getExcelVersionOfFilePath(gChatFilePath);
			saveWorkbook(outputExcelFilePath);
		}
		return outputExcelFilePath;
	}
	
	/**
	 * Processes the current line from the Google chat file, adding the line data to the spreadsheet
	 * in the necessary location
	 * 
	 * @param line
	 */
	private void processLine(String line) {
		if(line == null || line.trim().length() == 0 || workbook == null) return;
		// use first sheet
		
		boolean isMinutesDividerText = (runRegexOnString(line, MINUTES_DIVIDER_REGEX) != null);
		String timeText = runRegexOnString(line, TIME_REGEX);
		String nameText = runRegexOnString(line, NAME_REGEX);
		String text = runRegexOnString(line, TEXT_REGEX);
		
		if(timeText != null) {
			if((lastElementType == ChatElementType.TIME) || (lastElementType == ChatElementType.NAME) || (lastElementType == ChatElementType.TEXT)) {
				currentRow++;
			}
			Cell cell = setCell(currentRow, TIME_COLUMN_INDEX, timeText);
			cell.setCellStyle(TIME_CELL_STYLE);
			lastElementType = ChatElementType.TIME;
		}
		if(nameText != null) {
			if((lastElementType == ChatElementType.NAME) || (lastElementType == ChatElementType.TEXT)) {
				currentRow++;
			}
			lastName = nameText;
			Cell cell = setCell(currentRow, NAME_COLUMN_INDEX, nameText);
			cell.setCellStyle(NAME_CELL_STYLE);
			lastElementType = ChatElementType.NAME;
		}
		if(text != null) {
			if(lastElementType == ChatElementType.TEXT) {
				currentRow++;
			}
			Cell cell = setCell(currentRow, TEXT_COLUMN_INDEX, text);
			cell.setCellStyle(TEXT_CELL_STYLE);
			if(isMinutesDividerText) {
				cell.setCellStyle(MINUTES_DIVIDER_CELL_STYLE);
			}
			lastElementType = ChatElementType.TEXT;
		}
		
		// if no name exists, use the last name found
		String nameCellValue = getCellValue(currentRow, NAME_COLUMN_INDEX);
		if(nameCellValue == null || nameCellValue.trim().length() == 0 && lastName != null) {
			Cell cell = setCell(currentRow, NAME_COLUMN_INDEX, lastName);
			cell.setCellStyle(NAME_CELL_STYLE);
		}
	}
	
	/**
	 * Sets/creates a cell at the specified location with the supplied value
	 * 
	 * @param rowIndex
	 * @param columnIndex
	 * @param value
	 * @return
	 */
	private Cell setCell(int rowIndex, int columnIndex, String value) {
		Sheet sheet = workbook.getSheet("Chat");
		if(sheet == null) {
			sheet = workbook.createSheet("Chat");
			sheet.setColumnWidth(2, 20000);
		}
		Row row = sheet.getRow(rowIndex);
		if(row == null) {
			row = sheet.createRow(rowIndex);
		}
		Cell cell = row.getCell(columnIndex);
		if(cell == null) {
			cell = row.createCell(columnIndex);
		}
		CreationHelper createHelper = workbook.getCreationHelper();
		cell.setCellValue(createHelper.createRichTextString(value));
		
		return cell;
	}
	
	private String getCellValue(int rowIndex, int columnIndex) {
		Sheet sheet = workbook.getSheet("Chat");
		if(sheet != null) {
			Row row = sheet.getRow(rowIndex);
			if(row != null) {
				Cell cell = row.getCell(columnIndex);
				if(cell != null && cell.getCellType() == Cell.CELL_TYPE_STRING) {
					return cell.getRichStringCellValue().getString();
				}
			}
		}
		
		return null;
	}
	
	/**
	 * Creates workbook object
	 */
	private void initWorkbook() {
		this.workbook = new XSSFWorkbook();
		MINUTES_DIVIDER_CELL_STYLE = workbook.createCellStyle();
        Font redFont = workbook.createFont();
        redFont.setColor(IndexedColors.RED.getIndex());
        MINUTES_DIVIDER_CELL_STYLE.setFont(redFont);
        
        TIME_CELL_STYLE = workbook.createCellStyle();
        TIME_CELL_STYLE.setVerticalAlignment(CellStyle.VERTICAL_TOP);
        
        NAME_CELL_STYLE = workbook.createCellStyle();
        Font boldFont = workbook.createFont();
        boldFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        NAME_CELL_STYLE.setFont(boldFont);
        NAME_CELL_STYLE.setVerticalAlignment(CellStyle.VERTICAL_TOP);
        
        TEXT_CELL_STYLE = workbook.createCellStyle();
        TEXT_CELL_STYLE.setWrapText(true);
	}
	
	/**
	 * Write workbook to file
	 * 
	 * @param workbookFilePath
	 * @throws IOException
	 */
	private void saveWorkbook(String workbookFilePath) throws IOException {
		if(this.workbook != null) {
			FileOutputStream fileOut = new FileOutputStream(workbookFilePath);
			this.workbook.write(fileOut);
			fileOut.close();
		}
	}
	
	/**
	 * Returns the first regex group found in the supplied text using the supplied regex
	 * 
	 * @param text
	 * @param regex
	 * @return
	 */
	private String runRegexOnString(String text, String regex) {
		if(text == null || regex == null) return null;
		String value = null;
		Pattern pattern = Pattern.compile(regex);
		Matcher m = pattern.matcher(text);
		if(m.matches()) {
			if(m.groupCount() == 0) {
				value = m.group(0).trim();
			} else if (m.groupCount() > 0) {
				value = m.group(1).trim();
			}
		}
		if(value != null && value.trim().length() == 0) {
			value = null;
		}
		return value;
	}
	
	/**
	 * Returns supplied file path with it's extension replaced with an Excel file extension
	 * 
	 * @param filePath
	 * @return
	 */
	private String getExcelVersionOfFilePath(String filePath) {
		String result = null;
		if(filePath != null) {
			int lastPeriodIndex = filePath.lastIndexOf('.');
			if(lastPeriodIndex != -1 && lastPeriodIndex != filePath.length() - 1) {
				result = filePath.substring(0, filePath.lastIndexOf('.')) + ".xlsx";
			}
		}
		return result;
	}
	
	/**
	 * Checks if a file path has the specified extension
	 * 
	 * @param fileName
	 * @param extension
	 * @return
	 */
	private boolean hasSpecificExtension(String fileName, String extension) {
		if(fileName == null || extension == null) return false;
		return fileName.endsWith(extension);
	}
}
