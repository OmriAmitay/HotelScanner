package reports;

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
import org.apache.poi.hssf.usermodel.HSSFCell;
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
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import members.Currency;
import members.HotelEntity;
import members.Operation;
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
	public static final String LOGO_PATH = "c:\\sources\\HotelScanner\\reports\\logo.jpg";
	
	public static final int COLUMN_INDEX_CHECKIN_DATE = 0;
	public static final int COLUMN_INDEX_PRICE = 1;
	public static final int COLUMN_INDEX_HOTEL_NAME = 3;
	public static final int COLUMN_INDEX_CURRENCY = 4;
	
	private static final int PROVIDER_BUFFER = 3;
	
	public static int HEADLINE_ROW_INDEX = 0;
	public static int DATE_ROW_INDEX = 1;
	public static int DATA_ROW_INDEX = 2;
	public static int COMPETITORS_CELL_INDEX = 1;

	public static void main(String[] args) {

		try {

			// TODO: different parsers fullList, compertitors, date for each
			// provider: issta, daka90, eshet...
			// the report colors
			// add context for row and column index
			// S3 files download
			// send email

			// workbook, sheet and style
			XSSFWorkbook workbook = getWorkbook();
			XSSFSheet sheet = createSheet(workbook, SHEET_NAME_PREFIX + " " + getCurrentDate());
			Map<String, CellStyle> styles = createStyles(workbook);

			createHeadline(workbook, sheet, styles);
			createDatesHeadline(workbook, sheet, styles);
			createLogo(workbook, sheet);

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
						
						addRank(sheet, 12, length, Operation.RANK, styles);
						addOp(sheet, 13, length, Operation.AVERAGE, styles);
						addOp(sheet, 14, length, Operation.MIN, styles);
						addOp(sheet, 15, length, Operation.MAX, styles);
						
						DATA_ROW_INDEX = DATA_ROW_INDEX + length + 1 + PROVIDER_BUFFER;
						
						break;
						
					case ISSTA:
						
						System.out.println("Parsing Source="+source.toString());
						parser = new GenericParser(0, 1, 4, 5, ISSTA_SUBJECT_HOTEL, ISSTA_INPUT_DATA, true, "ISSTA");
						entites = readCSVInput(parser);
						
						competitorsNames = getCompetitorsNames(entites, parser.getSubjectHotel());
						System.out.println("Competitors " + competitorsNames);
						System.out.println("Found " + length + " compertitors, competitorsNames="+competitorsNames);
						
						length = competitorsNames.size();
						createProviderHeadline(workbook, sheet, styles, parser, length);
						createHotelLabels(workbook, sheet, styles, parser, competitorsNames);
						insertPricesByDates(workbook, sheet, styles, entites, length);
						
						addRank(sheet, 12, length, Operation.RANK, styles);
						addOp(sheet, 13, length, Operation.AVERAGE, styles);
						addOp(sheet, 14, length, Operation.MIN, styles);
						addOp(sheet, 15, length, Operation.MAX, styles);
						
						DATA_ROW_INDEX = DATA_ROW_INDEX + length + 1 + PROVIDER_BUFFER;
						
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
		sheet.addMergedRegion(new CellRangeAddress(DATA_ROW_INDEX, DATA_ROW_INDEX+length+4, 0, 0));
		XSSFRow yourRow = sheet.createRow(DATA_ROW_INDEX);
		XSSFCell providerCell = yourRow.createCell(0);
		providerCell.setCellValue(parser.getProviderName());
		providerCell.setCellStyle(styles.get("providerLabel"));
	}

	private static void insertPricesByDates(XSSFWorkbook workbook, XSSFSheet sheet, Map<String, CellStyle> styles, List<HotelEntity> entites, int competitoresLength) throws ParseException {

		for (int i = 2; i < 31; i++) {

			Date date = extractColumnDateFromComment(sheet, i);

			List<HotelEntity> filteredEntityByDate = entites.stream()
					.filter(e -> e.getCheckin() == date.getTime())
					.collect(Collectors.toList());

			if (! filteredEntityByDate.isEmpty()) {

				for (int k = DATA_ROW_INDEX; k < DATA_ROW_INDEX + competitoresLength + 1; k++) {

					XSSFRow competitrosRow = sheet.getRow(k);
					
					String competitorName = extractCompetitorName(sheet, competitrosRow);

					Optional<HotelEntity> competitorEntityByDate = filteredEntityByDate.stream()
							.filter(e -> e.getName().equals(competitorName))
							.findFirst();

					insertSinglePrice(workbook, sheet,styles, i, competitrosRow, competitorEntityByDate);

				}
			}
			
		}

	}
	
	private static void addOp(XSSFSheet sheet, int opRowIndex, int length, Operation operation, Map<String, CellStyle> styles) {
		XSSFRow rowFormula = sheet.createRow(opRowIndex);
		XSSFCell labelCell = rowFormula.createCell(COMPETITORS_CELL_INDEX);
		labelCell.setCellValue(operation.getTitle());
		labelCell.setCellStyle((styles.get("operationTitle")));
		
		for (int i = 2; i < 5; i++) {
			addAverageToColumn(sheet, rowFormula, i, length, operation, styles);
		}
		
	}
	
	private static void addAverageToColumn(XSSFSheet sheet, XSSFRow formulaRow, int columnIndex, int length, Operation operation, Map<String, CellStyle> styles) {
			XSSFRow row = sheet.getRow(DATA_ROW_INDEX);
			XSSFCell cell = row.getCell(columnIndex);
			String start = cell.getReference();
			row = sheet.getRow(DATA_ROW_INDEX + length);
			cell = row.getCell(columnIndex);
			String end = cell.getReference();
			String formula = operation.name() + "(" + start + ":" + end + ")";
			XSSFCell formulaCell = formulaRow.createCell(columnIndex);
			formulaCell.setCellType(HSSFCell.CELL_TYPE_FORMULA);
			formulaCell.setCellFormula(formula);
			formulaCell.setCellStyle((styles.get("operationValue")));
	}
	
	private static void addRank(XSSFSheet sheet, int opRowIndex, int length, Operation operation, Map<String, CellStyle> styles) {
		XSSFRow rowFormula = sheet.createRow(opRowIndex);
		XSSFCell labelCell = rowFormula.createCell(COMPETITORS_CELL_INDEX);
		labelCell.setCellValue(operation.getTitle());
		labelCell.setCellStyle((styles.get("operationTitle")));
		
		for (int i = 2; i < 5; i++) {
			addRankToColumn(sheet, rowFormula, i, length, operation, 1, styles);
		}
		
	}
	
	private static void addRankToColumn(XSSFSheet sheet, XSSFRow formulaRow, int columnIndex, int length, Operation operation, int sortType, Map<String, CellStyle> styles) {
			XSSFRow row = sheet.getRow(DATA_ROW_INDEX);
			XSSFCell cell = row.getCell(columnIndex);
			String start = cell.getReference();
			row = sheet.getRow(DATA_ROW_INDEX + length);
			cell = row.getCell(columnIndex);
			String end = cell.getReference();
			XSSFCell subjectCell = sheet.getRow(DATA_ROW_INDEX).getCell(columnIndex);
			
			String formula = operation.name() + "(" + subjectCell.getReference() + "," + start + ":" + end + "," + sortType + ")";
			XSSFCell formulaCell = formulaRow.createCell(columnIndex);
			formulaCell.setCellType(HSSFCell.CELL_TYPE_FORMULA);
			formulaCell.setCellFormula(formula);
			formulaCell.setCellStyle((styles.get("operationValue")));
	}
	
	private static String extractCompetitorName(XSSFSheet sheet, XSSFRow competitrosRow) {
		XSSFCell competitrosCell = competitrosRow.getCell(COMPETITORS_CELL_INDEX);
		return competitrosCell.getStringCellValue();
	}

	private static void insertSinglePrice(XSSFWorkbook workbook, XSSFSheet sheet,
										  Map<String, CellStyle> styles, int i, XSSFRow competitrosRow,
									      Optional<HotelEntity> competitorEntityByDate) {
		
		XSSFCell cell = competitrosRow.createCell(i);
		if(competitorEntityByDate.isPresent()) {
			cell.setCellValue(competitorEntityByDate.get().getPrice());
			//addCommentToCell(workbook, sheet, cell, "Comment");  // room type
			cell.setCellStyle(styles.get("available"));
		} else {
			cell.setCellValue("N/A");
			cell.setCellStyle(styles.get("notAvailable"));
		}
	}

	private static Date extractColumnDateFromComment(XSSFSheet sheet, int i) throws ParseException {
		XSSFRow row = sheet.getRow(DATES_ROW_INDEX);
		XSSFCell cell = row.getCell(i);
		XSSFComment cellComment = cell.getCellComment();
		XSSFRichTextString commentDate = cellComment.getString();
		DateFormat format = new SimpleDateFormat(REPORT_NORMALIZE_DATE_FORMAT);
		Date date = format.parse(commentDate.getString());
		return date;
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
		XSSFRow row = sheet.createRow( (short) HEADLINE_ROW_INDEX);
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
		XSSFRow yourRow = sheet.getRow(DATA_ROW_INDEX);
		XSSFCell yourCell = yourRow.createCell(COMPETITORS_CELL_INDEX);
		yourCell.setCellStyle(styles.get("subjectHotelName"));
		yourCell.setCellValue(parser.getSubjectHotel());

		// Competitors
		for (int i = 0; i < compertitors.size(); i++) {
			XSSFRow row = sheet.createRow(i + DATA_ROW_INDEX + 1);
			XSSFCell compertitorCell = row.createCell(COMPETITORS_CELL_INDEX);
			compertitorCell.setCellStyle(styles.get("hotelName"));
			compertitorCell.setCellValue(compertitors.get(i));
		}
		
	}

	private static void createDatesHeadline(XSSFWorkbook workbook, XSSFSheet sheet, Map<String, CellStyle> styles) {
		XSSFRow row = sheet.createRow(DATE_ROW_INDEX);
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

	private static void createLogo(XSSFWorkbook workbook, XSSFSheet sheet) throws IOException {
		createLogoPosition(workbook, sheet);
		
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

		short lightBoarderColor = IndexedColors.GREY_50_PERCENT.getIndex();
		short borderColor = IndexedColors.GREY_80_PERCENT.getIndex();

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
		providerFont.setFontName("Verdana");
		providerFont.setColor(HSSFColor.ROYAL_BLUE.index);
		providerFont.setItalic(true);
		providerFont.setBold(true);
		style = workbook.createCellStyle();
		style.setRotation((short) 90);
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setFont(providerFont);
		styles.put("providerLabel", style);

		XSSFFont dayFont = workbook.createFont();
		dayFont.setFontHeightInPoints((short) 10);
		style = workbook.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(CellStyle.ALIGN_CENTER);
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setRightBorderColor(lightBoarderColor);
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(lightBoarderColor);
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setLeftBorderColor(lightBoarderColor);
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setTopBorderColor(lightBoarderColor);
		style.setFont(dayFont);
		styles.put("workday", style);

		dayFont = workbook.createFont();
		dayFont.setFontHeightInPoints((short) 10);
		style = workbook.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(CellStyle.ALIGN_CENTER);
		style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
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
		XSSFFont font = workbook.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("Ariel");
		font.setColor(HSSFColor.ROYAL_BLUE.index);
		font.setBold(true);
		style = workbook.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setFont(font);
		styles.put("operationTitle", style);
		
		style = workbook.createCellStyle();
		font = workbook.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("Calibri");
		font.setColor(HSSFColor.ROYAL_BLUE.index);
		font.setBold(true);
		style = workbook.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setFont(font);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setRightBorderColor(lightBoarderColor);
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(lightBoarderColor);
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setLeftBorderColor(lightBoarderColor);
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setTopBorderColor(lightBoarderColor);
		styles.put("operationValue", style);
		//style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		//style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		
		font = workbook.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("Ariel");
		font.setBold(true);
		style = workbook.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setFont(font);
		styles.put("subjectHotelName", style);
		
		font = workbook.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("Ariel");
		style = workbook.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setFont(font);
		styles.put("hotelName", style);
		
		style = workbook.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setRightBorderColor(lightBoarderColor);
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(lightBoarderColor);
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setLeftBorderColor(lightBoarderColor);
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setTopBorderColor(lightBoarderColor);
		styles.put("available", style);

		style = workbook.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setRightBorderColor(lightBoarderColor);
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(lightBoarderColor);
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setLeftBorderColor(lightBoarderColor);
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setTopBorderColor(lightBoarderColor);
		style.setWrapText(true);
		styles.put("notAvailable", style);
		
		return styles;
	}
	
}
