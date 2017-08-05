package reports;

import java.awt.Color;
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.text.WordUtils;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import members.Currency;
import members.HotelEntity;
import members.Source;
import parsers.GenericParser;

public class ExcelUtils {

	public static final File output = new File("c:\\sources\\temp\\hotels-report.xlsx");

	public static final int DATES_ROW_INDEX = 1;
	public static final int COMPETITORS_ROW_START_INDEX = 2;

	public static Locale locale = new Locale("en");
	public static final String REPORT_NORMALIZE_DATE_FORMAT = "dd-MM-yyyy";
	public static final String SHEET_NAME_PREFIX = "Daily Report";

	public static String[] compertitors = { "שלמה המלך נתניה", "רזידנס נתניה", "איילנד נתניה", "השרון הרצליה", "דניאל הרצליה", "מלון קיו חוף פולג" };
	public static String[] issta_compertitors = { "המלך שלמה", "לאונרדו פלאזה נתניה", "איילנד", "רזידנס ביץ'", "רזידנס נתניה", "העונות" };

	public static final String ESHET_SUBJECT_HOTEL = "רמדה נתניה";
	public static final String ISSTA_SUBJECT_HOTEL = "רמדה";

	public static final String ESHET_INPUT_DATA = "c:\\sources\\HotelScanner\\reports\\eshettours\\inputData.xlsx";
	public static final String ISSTA_INPUT_DATA = "c:\\sources\\HotelScanner\\reports\\issta\\tester.xlsx";
	public static final String LOGO_PATH = "c:\\sources\\HotelScanner\\reports\\logic.png";
	
	public static final int COLUMN_INDEX_CHECKIN_DATE = 0;
	public static final int COLUMN_INDEX_PRICE = 1;
	public static final int COLUMN_INDEX_HOTEL_NAME = 3;
	public static final int COLUMN_INDEX_CURRENCY = 4;
	
	private static final int PROVIDER_BUFFER = 3;
	
	public static int headLineRowIndex = 0;
	public static int datesRowIndex = 1;
	public static int dataRowIndex = 2;

	public static void main(String[] args) {

		try {

			// TODO: different parsers fullList, compertitors, date for each
			// provider: issta, daka90, eshet...
			// the report colors
			// add context for row and column index
			// add avg for each room type + graph
			// S3 files download
			// send email

			XSSFWorkbook workbook = getWorkbook();
			XSSFSheet sheet = createSheet(workbook, SHEET_NAME_PREFIX + " " + getCurrentDate());
			Map<String, CellStyle> styles = createStyles(workbook);

			createHeadline(workbook, sheet, styles);
			createDatesHeadline(workbook, sheet, styles);
			
			createLogoPosition(workbook, sheet);
			setLogo(workbook, sheet);

			GenericParser parser = null;
			List<HotelEntity> entites = null;
			List<String> competitorsNames = null;
			int length = 0;
			
			List<Source> sourcesList = new ArrayList<>();
			//sourcesList.add(Source.ESHET);
			sourcesList.add(Source.ISSTA);
			for(Source source: sourcesList) {
			
				switch (source) {
				
					case ESHET:
						
						System.out.println("Parsing Source="+source.toString());	
						parser = new GenericParser(0, 1, 3, 4, ESHET_SUBJECT_HOTEL, ESHET_INPUT_DATA, true, "ESHET TOURS");
						entites = readCSVInput(parser);
						
						length = competitorsNames.size();
						competitorsNames = getCompetitorsNames(entites, parser.getSubjectHotel());
						System.out.println("Found " + length + " compertitors, competitorsNames="+competitorsNames);
						
						createProviderHeadline(workbook, sheet, styles, parser, length);
						createHotelLabels(workbook, sheet, styles, parser, competitorsNames);
						insertPricesByDates(workbook, sheet, styles, entites, length);
						
						dataRowIndex = dataRowIndex + length + 1 + PROVIDER_BUFFER;
						
						break;
						
					case ISSTA:
						
						parser = new GenericParser(0, 1, 4, 5, ISSTA_SUBJECT_HOTEL, ISSTA_INPUT_DATA, true, "ISSTA");
						entites = readCSVInput(parser);
						competitorsNames = getCompetitorsNames(entites, parser.getSubjectHotel());
						
						length = competitorsNames.size();
						createProviderHeadline(workbook, sheet, styles, parser, length);
						createHotelLabels(workbook, sheet, styles, parser, competitorsNames);
						insertPricesByDates(workbook, sheet, styles, entites, length);
						
						dataRowIndex = dataRowIndex + length + 1 + PROVIDER_BUFFER;
						
						break;
		
					default:
						break;
					}			
			}
			
			removeComments(workbook, sheet);
			closeFile(workbook, output);
			Desktop.getDesktop().open(output);

		} catch (Exception ex) {
			ex.printStackTrace();
		}

	}
	
	private static List<String> getCompetitorsNames(List<HotelEntity> entites, String subject) {
		Set<String> competitors = new HashSet<String>(); 
		entites.stream().forEach(e -> competitors.add(e.getName()));
		List<String> list = new ArrayList<String>(competitors);
		return list;
	}
	
	private static void removeComments(XSSFWorkbook workbook, XSSFSheet sheet) {
		XSSFRow row = sheet.getRow((short) 1);
		for (int i = 2; i < 33; i++) {
			XSSFCell cell = row.getCell(i);
			cell.removeCellComment();
		}

		/*
		 * for (int i = 2; i < 8; i++) { row = sheet.getRow((short) i); for(int
		 * k = 2; k < 20; k++) { XSSFCell cell = row.getCell(k);
		 * cell.removeCellComment(); } }
		 */

	}

	private static void createProviderHeadline(XSSFWorkbook workbook, XSSFSheet sheet, Map<String, CellStyle> styles, GenericParser parser, int length) {
		sheet.addMergedRegion(new CellRangeAddress(dataRowIndex, dataRowIndex+length, 0, 0));
		XSSFRow yourRow = sheet.createRow(dataRowIndex);
		XSSFCell providerCell = yourRow.createCell(0);
		providerCell.setCellValue(parser.getProviderName());
		providerCell.setCellStyle(styles.get("provider_label"));
	}

	private static void insertPricesByDates(XSSFWorkbook workbook, XSSFSheet sheet, Map<String, CellStyle> styles, List<HotelEntity> entites, int length) throws ParseException {

		for (int i = 2; i < 31; i++) {

			XSSFRow row = sheet.getRow(DATES_ROW_INDEX);
			XSSFCell cell = row.getCell(i);
			XSSFComment cellComment = cell.getCellComment();
			XSSFRichTextString commentDate = cellComment.getString();
			DateFormat format = new SimpleDateFormat(REPORT_NORMALIZE_DATE_FORMAT);
			Date date = format.parse(commentDate.getString());

			List<HotelEntity> filteredByDate = entites.stream()
					.filter(e -> e.getCheckin() == date.getTime())
					.collect(Collectors.toList());

			if (! filteredByDate.isEmpty()) {

				for (int k = dataRowIndex; k < dataRowIndex + length + 1; k++) {

					XSSFRow competitrosRow = sheet.getRow(k);
					XSSFCell competitrosCell = competitrosRow.getCell(1);
					String competitorName = competitrosCell.getStringCellValue();

					Optional<HotelEntity> hotelEntityByDate = filteredByDate.stream()
							.filter(e -> e.getName().equals(competitorName))
							.findFirst();

					cell = competitrosRow.createCell(i);
					if(hotelEntityByDate.isPresent()) {
						cell.setCellValue(hotelEntityByDate.get().getPrice());
						cell.setCellStyle(styles.get("available"));
					} else {
						cell.setCellValue("N/A");
						cell.setCellStyle(styles.get("not available"));
					}

				}
			}
		}

	}

	private static XSSFWorkbook getWorkbook() throws IOException {
		XSSFWorkbook workbook = new XSSFWorkbook();
		return workbook;
	}

	private static void closeFile(XSSFWorkbook workbook, File output) throws IOException {
		// Create file system using specific name
		FileOutputStream out = new FileOutputStream(output);
		// write operation workbook using file out object
		workbook.write(out);
		out.close();
		System.out.println("createworkbook.xlsx written successfully");
	}

	private static XSSFSheet createSheet(XSSFWorkbook workbook, String sheetName) {
		return workbook.createSheet(sheetName);
	}

	private static void createHeadline(XSSFWorkbook workbook, XSSFSheet sheet, Map<String, CellStyle> styles) {
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 2, 33));
		XSSFRow row = sheet.createRow( (short) headLineRowIndex);
		XSSFCell cell = row.createCell(2);
		cell.setCellValue("Daily Report");
		cell.setCellStyle(styles.get("headline"));
	}

	private static String getCurrentDate() {
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

	private static void createLogoPosition(XSSFWorkbook workbook, XSSFSheet sheet) {
		sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 1));
		sheet.setColumnWidth(0, 1000);
		sheet.setColumnWidth(1, 6000);
	}

	private static void createHotelLabels(XSSFWorkbook workbook, XSSFSheet sheet, Map<String, CellStyle> styles,
										  GenericParser parser, List<String> compertitors) {

		// Subject
		XSSFRow yourRow = sheet.getRow(dataRowIndex);
		XSSFCell yourCell = yourRow.createCell(1);
		yourCell.setCellStyle(styles.get("subject hotel name"));
		yourCell.setCellValue(parser.getSubjectHotel());

		for (int i = 0; i < compertitors.size(); i++) {
			XSSFRow row = sheet.createRow(i + dataRowIndex + 1);
			XSSFCell compertitorCell = row.createCell(1);
			compertitorCell.setCellStyle(styles.get("hotel name"));
			compertitorCell.setCellValue(compertitors.get(i));
		}
		
	}

	private static void createDatesHeadline(XSSFWorkbook workbook, XSSFSheet sheet, Map<String, CellStyle> styles) {
		XSSFRow row = sheet.createRow(datesRowIndex);
		row.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));
		//Calendar calendar = Calendar.getInstance();
		//calendar.setTime(new Date());
		Calendar calendar = new GregorianCalendar(2017, 7, 19, 12, 00, 00);

		CellStyle firstStyle = workbook.createCellStyle();
		firstStyle.setAlignment(CellStyle.ALIGN_CENTER);
		XSSFCell firstCell = row.createCell(1);
		firstCell.setCellStyle(firstStyle);
		firstCell.setCellValue("\n\n\n");

		for (int i = 2; i < 33; i++) {

			calendar.add(Calendar.DATE, 1);
			int dayOfWeek = calendar.get(Calendar.DAY_OF_WEEK);
			Date date = calendar.getTime();
			SimpleDateFormat simpleDateFormat = new SimpleDateFormat("EEE");
			String dayStr = simpleDateFormat.format(date).toUpperCase();
			dayStr = camelCase(dayStr);
			XSSFCell cell = row.createCell(i);

			if (dayOfWeek == 6 || dayOfWeek == 7) {
				cell.setCellStyle(styles.get("weekend"));
			} else {
				cell.setCellStyle(styles.get("workday"));
			}

			cell.setCellValue(calendar.getDisplayName(Calendar.MONTH, Calendar.SHORT_FORMAT, locale) + "-" + calendar.get(Calendar.DAY_OF_MONTH));

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

	private static String camelCase(String text) {
		return StringUtils.remove(WordUtils.capitalizeFully(text, '_'), "_");
	}

	private static List<HotelEntity> readCSVInput(GenericParser parser) {

		List<HotelEntity> entites = new LinkedList<>();
		try {
			File myFile = new File(parser.getInput());
			FileInputStream fis = new FileInputStream(myFile);
			XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);
			XSSFSheet mySheet = myWorkBook.getSheetAt(0);
			
			Iterator<Row> rowIterator = mySheet.iterator();
			if(parser.isSkipHeadline()) {
				rowIterator.next();
			}

			while (rowIterator.hasNext()) {
				HotelEntity hotelEntity = new HotelEntity();
				Row row = rowIterator.next();

				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();

					int columnIndex = cell.getColumnIndex();
					if (columnIndex == parser.getChekcinDateIndex()) {
						Date date = parser.parseDateField(cell, REPORT_NORMALIZE_DATE_FORMAT);
						hotelEntity.setCheckin(date != null ? date.getTime() : null);
					}

					if (columnIndex == parser.getHotelNameIndex()) {
						hotelEntity.setName(parser.parseStringField(cell));
					}

					if (columnIndex == parser.getPriceIndex()) {
						hotelEntity.setPrice(parser.parseNumericField(cell));
					}

					if (columnIndex == parser.getCurrencyIndex()) {
						hotelEntity.setCurrency(parser.parseStringField(cell).equals("₪") ? Currency.NIS : Currency.DOLLAR);
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

	private static void setLogo(XSSFWorkbook workbook, XSSFSheet sheet) throws IOException {
		InputStream inputStream = new FileInputStream(LOGO_PATH);
		byte[] imageBytes = IOUtils.toByteArray(inputStream);
		int pictureureIdx = workbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);
		inputStream.close();
		CreationHelper helper = workbook.getCreationHelper();
		Drawing drawing = sheet.createDrawingPatriarch();
		ClientAnchor anchor = helper.createClientAnchor();
		anchor.setRow1(0);
		anchor.setCol1(1);
		Picture picture = drawing.createPicture(anchor, pictureureIdx);
		picture.resize(0.9d);
	}

	/**
	 * cell styles used for formatting calendar sheets
	 */
	private static Map<String, CellStyle> createStyles(XSSFWorkbook workbook) {
		Map<String, CellStyle> styles = new HashMap<String, CellStyle>();

		short borderColor = IndexedColors.GREY_50_PERCENT.getIndex();

		CellStyle style = workbook.createCellStyle();
		
		Font titleFont = workbook.createFont();
		titleFont.setFontHeightInPoints((short) 26);
		titleFont.setColor(IndexedColors.DARK_BLUE.getIndex());
		titleFont.setFontName("Bookman Old Style");
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(CellStyle.ALIGN_CENTER);
		style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
		style.setFont(titleFont);
		styles.put("headline", style);
		
		XSSFFont providerFont = workbook.createFont();
		providerFont.setFontHeightInPoints((short) 14);
		providerFont.setFontName("Ariel");
		providerFont.setColor(HSSFColor.ROYAL_BLUE.index);
		providerFont.setItalic(true);
		providerFont.setBold(true);
		style = workbook.createCellStyle();
		style.setRotation((short) 90);
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setFont(providerFont);
		styles.put("provider_label", style);

		/*XSSFFont monthFont = workbook.createFont();
		monthFont.setFontHeightInPoints((short) 12);
		monthFont.setColor(IndexedColors.WHITE.getIndex());
		monthFont.setBold(true);
		style = workbook.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(CellStyle.ALIGN_CENTER);
		style.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setFont(monthFont);
		styles.put("month", style);*/

		XSSFFont dayFont = workbook.createFont();
		dayFont.setFontHeightInPoints((short) 10);
		//dayFont.setBold(true);
		style = workbook.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(CellStyle.ALIGN_CENTER);
		style.setFillForegroundColor(IndexedColors.SKY_BLUE.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setRightBorderColor(borderColor);
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(borderColor);
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setLeftBorderColor(borderColor);
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setTopBorderColor(borderColor);
		style.setFont(dayFont);
		styles.put("workday", style);

		dayFont = workbook.createFont();
		dayFont.setFontHeightInPoints((short) 10);
		//dayFont.setBold(true);
		style = workbook.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(CellStyle.ALIGN_CENTER);
		style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setRightBorderColor(borderColor);
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(borderColor);
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setLeftBorderColor(borderColor);
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setTopBorderColor(borderColor);
		style.setFont(dayFont);
		styles.put("weekend", style);
		
		style = workbook.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_JUSTIFY);
		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_JUSTIFY);
		//style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		//style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setRightBorderColor(borderColor);
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(borderColor);
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setLeftBorderColor(borderColor);
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setTopBorderColor(borderColor);
		styles.put("available", style);
		
		XSSFFont font = workbook.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("Ariel");
		//font.setColor(HSSFColor.ROYAL_BLUE.index);
		//font.setItalic(true);
		//font.setBold(true);
		//font.setFontName("Segoe Print");
		style = workbook.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setFont(font);
		styles.put("hotel name", style);
		
		font = workbook.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("Ariel");
		font.setColor(HSSFColor.ROYAL_BLUE.index);
		font.setBold(true);
		style = workbook.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setFont(font);
		styles.put("subject hotel name", style);

		/*style = workbook.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(CellStyle.VERTICAL_TOP);
		style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setRightBorderColor(borderColor);
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(borderColor);
		styles.put("weekend_right", style);

		style = workbook.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_LEFT);
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		style.setVerticalAlignment(CellStyle.VERTICAL_TOP);
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setLeftBorderColor(borderColor);
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(borderColor);
		style.setFont(dayFont);
		styles.put("workday_left", style);*/

		/*style = workbook.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(CellStyle.VERTICAL_TOP);
		style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setRightBorderColor(borderColor);
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(borderColor);
		styles.put("workday_right", style);

		style = workbook.createCellStyle();
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(borderColor);
		styles.put("grey_left", style);

		style = workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setRightBorderColor(borderColor);
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(borderColor);
		styles.put("grey_right", style);
*/
		

		style = workbook.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		//style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		//style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setRightBorderColor(borderColor);
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(borderColor);
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setLeftBorderColor(borderColor);
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setTopBorderColor(borderColor);
		styles.put("available", style);

		style = workbook.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		//style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
		//style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setRightBorderColor(borderColor);
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(borderColor);
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setLeftBorderColor(borderColor);
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setTopBorderColor(borderColor);
		style.setWrapText(true);
		styles.put("not available", style);
		
		style = workbook.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setFont(font);
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		styles.put("hotel_name_second", style);

		return styles;
	}
	
	/*public HSSFColor setColor(XSSFWorkbook sheet, byte r,byte g, byte b){
	    HSSFPalette palette = sheet.getCustomPalette();
	    HSSFColor hssfColor = null;
	    try {
	        hssfColor= palette.findColor(r, g, b); 
	        if (hssfColor == null ){
	            palette.setColorAtIndex(HSSFColor.LAVENDER.index, r, g,b);
	            hssfColor = palette.getColor(HSSFColor.LAVENDER.index);
	        }
	    } catch (Exception e) {
	        //logger.error(e);
	    }

	    return hssfColor;
	}*/

}
