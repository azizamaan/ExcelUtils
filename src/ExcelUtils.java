/*Excel Reader is a utility class which reads excel and converts to List of VOs and vice versa. It also validates header row.
 * 
 */

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.multipart.MultipartFile;

/**
 * @author amaan
 *
 */
public class ExcelUtils {
	private static final String EXCEPTION_OCCURED = "Exception occured: ";
	private static Logger logger = LoggerFactory.getLogger(ExcelUtils.class);
	private static final String MM_DD_YYYY = "MM/dd/yyyy";

	/**
	 * ExcelConverter reads the file provided in parameters and converts to list of objects of class type provided in
	 * parameters.
	 * 
	 * @param file
	 * @param className
	 * @return
	 */
	public static <T> List<T> excelConverter(MultipartFile file, Class<T> className) {
		Workbook workbook = null;
		try {
			workbook = WorkbookFactory.create(file.getInputStream());
		} catch (InvalidFormatException e) {
			logger.debug(EXCEPTION_OCCURED + e);
		} catch (IOException e) {
			logger.debug(EXCEPTION_OCCURED + e);
		}
		Sheet worksheet = workbook.getSheetAt(0);
		Row headerRow = worksheet.getRow(worksheet.getFirstRowNum());
		
		List<T> list = new ArrayList<T>();
		for (Row row : worksheet) {
			if (headerRow != row) {
				list.add(className.cast(excelRowMapper(row, headerRow, className)));
			}
		}
		return list;
	}

	/**
	 * This method converts an excel row to corresponding object.
	 * 
	 * @param currentRow
	 * @param headerRow
	 * @param className
	 * @return
	 */
	public static Object excelRowMapper(Row currentRow, Row headerRow, Class className) {
		Map<String, String> excelRowMap = new HashMap<String, String>();
		for (Cell cell : currentRow) {
			excelRowMap.put(getCamelCaseString(getExcelCellValue(headerRow.getCell(cell.getColumnIndex()))), getExcelCellValue(cell));
		}
		Object object = null;
		try {
			object = className.newInstance();
			BeanUtils.populate(object, excelRowMap);
		} catch (InstantiationException e) {
			logger.debug(EXCEPTION_OCCURED + e);
		} catch (IllegalAccessException e) {
			logger.debug(EXCEPTION_OCCURED + e);
		} catch (InvocationTargetException e) {
			logger.debug(EXCEPTION_OCCURED + e);
		}
		return object;
	}

	/**
	 * This method returns string value of cell.
	 * 
	 * @param cell
	 * @return
	 */
	public static String getExcelCellValue(Cell cell) {
		DataFormatter df = new DataFormatter();
		switch (cell.getCellType()) {
			case Cell.CELL_TYPE_STRING:
				return cell.getRichStringCellValue().getString();
			case Cell.CELL_TYPE_NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					return new SimpleDateFormat(MM_DD_YYYY).format(cell.getDateCellValue());
				} else {
					String num = df.formatCellValue(cell);
					return !num.contains(".") ? num : num.replaceAll("0*$", "").replaceAll("\\.$", "");
				}
			case Cell.CELL_TYPE_BOOLEAN:
				return "" + cell.getBooleanCellValue();
			case Cell.CELL_TYPE_FORMULA:
				return cell.getCellFormula();
			case Cell.CELL_TYPE_BLANK:
				return "";
			case Cell.CELL_TYPE_ERROR:
				return Byte.valueOf(cell.getErrorCellValue()).toString();
			default:
				return "";
		}
	}

	/**
	 * This method converts underscored or spaced string to it's camel case format.
	 * 
	 * @param string
	 * @return
	 */
	public static String getCamelCaseString(String string) {
		StringBuffer sb = new StringBuffer();
		if(null != string && !"".equals(string)){
			for (String s : string.split(" |_|-|/")) {
				sb.append(Character.toUpperCase(s.charAt(0)));
				if (s.length() > 1) {
					sb.append(s.substring(1, s.length()).toLowerCase());
				}
			}
			sb.setCharAt(0, Character.toLowerCase(sb.charAt(0)));
		}
		return sb.toString().replaceAll("[^A-Za-z0-9]", "");
	}
	
	
	/**
	 * Generates an excel report from list of objects and header column.
	 * @param <T>
	 * 
	 * @param failedSheet
	 * @return
	 */
	public static <T> File generateExcel(Map<String, List<T>> listMap, Map<String, ArrayList<String>> headerMap, Map<String, String> localeLabels) {
		HSSFWorkbook wb = new HSSFWorkbook();
		
		for (Map.Entry<String, ArrayList<String>> entry : headerMap.entrySet()) {
		    String sheetName = entry.getKey();
		    ArrayList<String> header = entry.getValue();
		    
		    HSSFSheet sheet = wb.createSheet(localeLabels.get(sheetName));

		    int cellIndex = 0;
		    int rowIndex = 0;
	    
		    HSSFRow headerRow = sheet.createRow(rowIndex++);
	    
		    //Cell styles: Header Cell
		    HSSFCellStyle headerStyle = (HSSFCellStyle) wb.createCellStyle();
		    HSSFPalette palette = wb.getCustomPalette();
		    palette.setColorAtIndex(HSSFColor.BLUE.index, (byte) 47, (byte) 117, (byte) 181);
		    headerStyle.setFillForegroundColor(HSSFColor.BLUE.index);
		    headerStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		 
			HSSFFont headerFont = (HSSFFont) wb.createFont();
			//headerFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
			headerStyle.setFont(headerFont);
			headerFont.setColor(IndexedColors.WHITE.getIndex());
		 
			headerStyle.setBorderBottom(CellStyle.BORDER_THIN);
			headerStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
			headerStyle.setBorderLeft(CellStyle.BORDER_THIN);
			headerStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
			headerStyle.setBorderRight(CellStyle.BORDER_THIN);
			headerStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
			headerStyle.setBorderTop(CellStyle.BORDER_THIN);
			headerStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
			
			headerStyle.setAlignment(CellStyle.ALIGN_CENTER);
			
			//Cell styles: Body Cell
			HSSFCellStyle style = (HSSFCellStyle) wb.createCellStyle();
		 
			style.setBorderBottom(CellStyle.BORDER_THIN);
			style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
			style.setBorderLeft(CellStyle.BORDER_THIN);
			style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
			style.setBorderRight(CellStyle.BORDER_THIN);
			style.setRightBorderColor(IndexedColors.BLACK.getIndex());
			style.setBorderTop(CellStyle.BORDER_THIN);
			style.setTopBorderColor(IndexedColors.BLACK.getIndex());
			
			style.setAlignment(CellStyle.ALIGN_CENTER);
			style.setWrapText(true);
			
		
			for (Iterator iterator = header.iterator(); iterator.hasNext();) {
				HSSFCell cell = headerRow.createCell(cellIndex); 
				HSSFRichTextString string = new HSSFRichTextString(localeLabels.get((String) iterator.next()));
				cell.setCellValue(string);
				
				cell.setCellStyle(headerStyle);
				sheet.autoSizeColumn((short) cellIndex++);
			}
		
		
			List<Map<Object, Object>> beanMapList = new ArrayList<Map<Object, Object>>();
			List<Object> list = (List<Object>) listMap.get(sheetName);
			for (Iterator<Object> iterator = list.iterator(); iterator.hasNext();) {
				Object object = iterator.next();
				Map<Object, Object> beanMap = new org.apache.commons.beanutils.BeanMap(object);
				beanMapList.add(beanMap);
				
			}

			for (Iterator<Map<Object, Object>> iterator = beanMapList.iterator(); iterator.hasNext();) {
				HSSFRow row = sheet.createRow(rowIndex++);
				cellIndex = 0;
				Map<Object, Object> beanMap = iterator.next();
				for (Iterator<String> headerIterator = header.iterator(); headerIterator.hasNext();) {
					String headerKey = headerIterator.next();
					String key = getCamelCaseString(headerKey);
					HSSFRichTextString value = new HSSFRichTextString(null != (String) beanMap.get(key) ? (String) beanMap.get(key) : "");
					
					HSSFCell cell = row.createCell(cellIndex++); 
					cell.setCellValue(value);
					cell.setCellStyle(style);
					
				}
			}
		}

		// Write the output to a file
		File file = null;
	    FileOutputStream fileOut = null;
		try {
			file = new File("ExcelFile.xls");
			//file = new File("D:\\ExcelFile.xls");
			fileOut = new FileOutputStream(file);
			wb.write(fileOut);
		    fileOut.close();
		} catch (IOException e) {
			logger.debug(EXCEPTION_OCCURED + e);
		}
	    
		return file;
	}
	
	/**
	 * ValidateHeaderRow reads the file provided in parameters and validates header row to list of names provided in
	 * parameters.
	 * 
	 * @param file
	 * @param className
	 * @return
	 */
	public static boolean validateHeaderRow(MultipartFile file, ArrayList<String> headerRowNames) {
		Workbook workbook = null;
		try {
			workbook = WorkbookFactory.create(file.getInputStream());
		} catch (InvalidFormatException e) {
			logger.debug(EXCEPTION_OCCURED + e);
		} catch (IOException e) {
			logger.debug(EXCEPTION_OCCURED + e);
		}
		Sheet worksheet = workbook.getSheetAt(0);
		Row headerRow = worksheet.getRow(worksheet.getFirstRowNum());
		
		for (Cell cell : headerRow) {
			String cellValue = getExcelCellValue(cell);
			if(null != cellValue && !"".equals(cellValue)){
				String columnValue = getCamelCaseString(cellValue);
				if(null != columnValue && !"".equals(columnValue)){
					Boolean matchFlag = false;
					for (Iterator iterator = headerRowNames.iterator(); iterator.hasNext();) {
						String headerValue = getCamelCaseString((String) iterator.next());
						if(columnValue.equals(headerValue)){
							matchFlag = true;
							break;
						}
					}
					if(!matchFlag){
						return false;
					}
				}
			}
		}
		return true;
	}
}
