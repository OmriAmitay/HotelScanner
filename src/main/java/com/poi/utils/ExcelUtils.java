package com.poi.utils;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.text.WordUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import parser.Currency;
import parser.HotelEntity;

public class ExcelUtils {

	public static final File output = new File("c:\\sources\\temp\\hotels-report.xlsx");
	public static String[] compertitors = { "שלמה המלך נתניה", "רזידנס נתניה", "איילנד נתניה", "השרון הרצליה",
			"דניאל הרצליה", "מלון קיו חוף פולג" };
	public static String[] fullList = { "שלמה המלך נתניה", "רזידנס נתניה", "איילנד נתניה", "השרון הרצליה",
			"דניאל הרצליה", "מלון קיו חוף פולג", "רמדה נתניה" };
	public static final String SUBJECT_HOTEL = "רמדה נתניה";

	public static final int DATES_ROW_INDEX = 1;
	public static final int COMPETITORS_ROW_START_INDEX = 2;

	public static final String REPORT_NORMALIZE_DATE_FORMAT = "dd-MM-yyyy";

	public static final String INPUT_DATA = "c:\\sources\\temp\\inputData.xlsx";
	public static final String INPUT_DATA_TEST = "c:\\sources\\temp\\1.xlsx";

	public static void main(String[] args) {

		try {

			XSSFWorkbook workbook = getWorkbook();

			XSSFSheet sheet = createSheet(workbook, "Daily Report " + getCurrentDate());

			createHeadline(workbook, sheet);

			createDatesHeadline(workbook, sheet);

			createHotelLabels(workbook, sheet);

			List<HotelEntity> entites = readCSVInput(INPUT_DATA);

			insertPricesByDates(workbook, sheet, entites);

			closeFile(workbook, output);
			Desktop.getDesktop().open(output);

		} catch (Exception ex) {
			ex.printStackTrace();
		}

	}

	private static void insertPricesByDates(XSSFWorkbook workbook, XSSFSheet sheet, List<HotelEntity> entites)
			throws ParseException {

		for (int i = 1; i < 31; i++) {

			XSSFRow row = sheet.getRow(DATES_ROW_INDEX);
			XSSFCell cell = row.getCell(i);
			XSSFComment cellComment = cell.getCellComment();
			XSSFRichTextString commentDate = cellComment.getString();
			DateFormat format = new SimpleDateFormat(REPORT_NORMALIZE_DATE_FORMAT);
			Date date = format.parse(commentDate.getString());

			List<HotelEntity> filteredByExactDate = entites.stream().filter(e -> e.getCheckin() == date.getTime())
					.collect(Collectors.toList());

			if (!filteredByExactDate.isEmpty()) {

				for (int k = 2; k < 8; k++) {

					XSSFRow competitrosRow = sheet.getRow(k);
					XSSFCell competitrosCell = competitrosRow.getCell(0);
					String competitorName = competitrosCell.getStringCellValue();

					Optional<HotelEntity> hotelEntity = filteredByExactDate.stream()
							.filter(e -> e.getName().equals(competitorName)).findFirst();

					row = sheet.getRow(k);
					cell = row.createCell(i);
					cell.setCellValue(hotelEntity.isPresent() ? Double.toString(hotelEntity.get().getPrice()) : "N/A");

					if (!hotelEntity.isPresent()) {
						CellStyle yellow = workbook.createCellStyle();
						yellow.setAlignment(CellStyle.ALIGN_CENTER);
						yellow.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
						yellow.setFillPattern(CellStyle.SOLID_FOREGROUND);
						yellow.setWrapText(true);
						cell.setCellStyle(yellow);
					}

				}
			}
		}

	}

	public static XSSFWorkbook getWorkbook() throws IOException {
		XSSFWorkbook workbook = new XSSFWorkbook();
		return workbook;
	}

	public static void closeFile(XSSFWorkbook workbook, File output) throws IOException {
		// Create file system using specific name
		FileOutputStream out = new FileOutputStream(output);
		// write operation workbook using file out object
		workbook.write(out);
		out.close();
		System.out.println("createworkbook.xlsx written successfully");
	}

	public static XSSFSheet createSheet(XSSFWorkbook workbook, String sheetName) {
		return workbook.createSheet(sheetName);
	}

	public static void createHeadline(XSSFWorkbook workbook, XSSFSheet sheet) {
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 30));

		XSSFRow row = sheet.createRow((short) 0);
		XSSFCell cell = row.createCell(0);
		cell.setCellValue("Daily Report");

		XSSFFont font = workbook.createFont();
		font.setFontHeightInPoints((short) 30);
		font.setFontName("IMPACT");
		font.setItalic(true);
		font.setColor(HSSFColor.RED.index);

		XSSFCellStyle style = workbook.createCellStyle();
		style.setFont(font);
		style.setFillForegroundColor(HSSFColor.ROYAL_BLUE.index);
		style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

		cell.setCellStyle(style);
	}

	public static String getCurrentDate() {
		SimpleDateFormat sdfDate = new SimpleDateFormat("dd-MM-yy");
		Date now = new Date();
		String date = sdfDate.format(now);
		return date;
	}

	private static String convertTimestampToDateAsString(long timeStamp) {
		String date = null;
		try {
			DateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
			Date netDate = (new Date(timeStamp));
			date = sdf.format(netDate);
		} catch (Exception ex) {
			ex.printStackTrace();
		}

		return date;
	}

	private static void createHotelLabels(XSSFWorkbook workbook, XSSFSheet sheet) {
		XSSFFont font = workbook.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("Ariel");
		font.setColor(HSSFColor.ROYAL_BLUE.index);
		font.setItalic(true);
		font.setBold(true);

		sheet.setColumnWidth(0, 5000);
		CellStyle style = workbook.createCellStyle();
		XSSFRow yourRow = sheet.createRow((short) 2);
		XSSFCell yourCell = yourRow.createCell(0);
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setFont(font);
		yourCell.setCellStyle(style);
		yourCell.setCellValue(SUBJECT_HOTEL);

		for (int i = 0; i < 5; i++) {

			XSSFRow row = sheet.createRow((short) i + 3);
			XSSFCell compertitorCell = row.createCell(0);

			CellStyle competitorStyle = workbook.createCellStyle();
			competitorStyle.setAlignment(CellStyle.ALIGN_CENTER);
			competitorStyle.setFont(font);

			if (i % 2 == 0) {
				competitorStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
				competitorStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
			}

			compertitorCell.setCellStyle(competitorStyle);
			compertitorCell.setCellValue(compertitors[i]);

			addCommentToCell(workbook, sheet, compertitorCell, Integer.valueOf(i).toString());
		}
	}

	private static void createDatesHeadline(XSSFWorkbook workbook, XSSFSheet sheet) {
		XSSFRow row = sheet.createRow((short) 1);
		row.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));
		Calendar calendar = Calendar.getInstance();
		calendar.setTime(new Date());

		CellStyle center = workbook.createCellStyle();
		center.setAlignment(CellStyle.ALIGN_CENTER);
		center.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
		center.setFillPattern(CellStyle.SOLID_FOREGROUND);
		center.setWrapText(true);

		Font font = workbook.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("Ariel");
		font.setItalic(true);
		center.setFont(font);

		CellStyle firstStyle = workbook.createCellStyle();
		firstStyle.setAlignment(CellStyle.ALIGN_CENTER);

		XSSFCell firstCell = row.createCell(0);
		firstCell.setCellStyle(firstStyle);
		firstCell.setCellValue("\n\n\n");

		for (int i = 1; i < 31; i++) {

			calendar.add(Calendar.DATE, 1);
			int dayOfMonth = calendar.get(Calendar.DAY_OF_MONTH);
			Date date = calendar.getTime();
			SimpleDateFormat simpleDateFormat = new SimpleDateFormat("EEE");
			String dayStr = simpleDateFormat.format(date).toUpperCase();
			dayStr = camelCase(dayStr);
			XSSFCell cell = row.createCell(i);
			cell.setCellStyle(center);
			cell.setCellValue(dayStr + "\n" + dayOfMonth);

			simpleDateFormat = new SimpleDateFormat("dd-MM-yyyy");
			dayStr = simpleDateFormat.format(date).toUpperCase();

			addCommentToCell(workbook, sheet, cell, dayStr); // add full date
																// for search
		}
	}

	private static void addCommentToCell(XSSFWorkbook workbook, XSSFSheet sheet, XSSFCell cell, String value) {
		SimpleDateFormat simpleDateFormat;
		String dayStr;
		CreationHelper factory = workbook.getCreationHelper();
		Drawing drawing = sheet.createDrawingPatriarch();
		ClientAnchor anchor = factory.createClientAnchor();
		Comment comment = drawing.createCellComment(anchor);
		RichTextString str = factory.createRichTextString(value);
		comment.setString(str);
		comment.setAuthor("Apache POI");
		cell.setCellComment(comment);
	}

	public static String camelCase(String text) {
		return StringUtils.remove(WordUtils.capitalizeFully(text, '_'), "_");
	}

	public static List<HotelEntity> readCSVInput(String input) {

		List<HotelEntity> entites = new LinkedList<>();
		try {

			File myFile = new File(input);
			FileInputStream fis = new FileInputStream(myFile);
			XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);
			XSSFSheet mySheet = myWorkBook.getSheetAt(0);

			// ignore headlines
			Iterator<Row> rowIterator = mySheet.iterator();
			Row row = rowIterator.next();

			while (rowIterator.hasNext()) {

				HotelEntity hotelEntity = new HotelEntity();
				row = rowIterator.next();

				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) {

					Cell cell = cellIterator.next();

					int columnIndex = cell.getColumnIndex();
					int cellType = cell.getCellType();

					if (columnIndex == 0) {

						DateFormat format = new SimpleDateFormat(REPORT_NORMALIZE_DATE_FORMAT);

						Date date = null;
						String dateStr = null;
						if (cellType == 1) {
							dateStr = cell.getStringCellValue();

							try {
								date = format.parse(dateStr);
							} catch (Exception ex) {
								String[] split = dateStr.split("/");
								if (split[0].length() > 2) {
									split[0] = split[0].substring(1, split[0].length());
								} else if (Integer.valueOf(split[0]) < 10) {
									split[0] = "0" + split[0];
								}

								dateStr = split[0] + "-" + split[1] + "-" + split[2];
								date = format.parse(dateStr);
							}

						} else if (cellType == 0) {
							date = cell.getDateCellValue();
						}

						hotelEntity.setCheckin(date != null ? date.getTime() : null);
					}

					if (columnIndex == 1) {
						hotelEntity.setName(cell.getStringCellValue());
					}

					if (columnIndex == 3) {
						hotelEntity.setPrice(cell.getNumericCellValue());
					}

					if (columnIndex == 4) {
						hotelEntity.setCurrency(cell.getStringCellValue().equals("₪") ? Currency.NIS : Currency.DOLLAR);
					}

				}

				entites.add(hotelEntity);

			}
		} catch (Exception ex) {
			System.err.println(ex);
		}

		/*
		 * Collections.sort(entites, new Comparator<HotelEntity>() {
		 * 
		 * @Override public int compare(HotelEntity o1, HotelEntity o2) { return
		 * (int) (o1.getCheckin() - o2.getCheckin()); } });
		 */

		return entites;

	}

}
