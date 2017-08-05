package parsers;

import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;

public class GenericParser implements CSVParser {
	
	public int chekcinDateIndex;
	public int priceIndex;
	public int hotelNameIndex;
	public int currencyIndex;
	public String subjectHotel;
	public boolean skipHeadline;
	public String input;
	public String providerName;
	
	public GenericParser(int chekcinDateIndex, int hotelNameIndex, int priceIndex, int currencyIndex, String subjectHotel, String input, boolean skipHeadline, String providerName) {
		this.chekcinDateIndex = chekcinDateIndex;
		this.priceIndex = priceIndex;
		this.hotelNameIndex = hotelNameIndex;
		this.currencyIndex = currencyIndex;
		this.subjectHotel = subjectHotel;
		this.input = input;
		this.skipHeadline = skipHeadline;
		this.providerName = providerName;
	}
	
	public int getChekcinDateIndex() {
		return chekcinDateIndex;
	}
	
	public void setChekcinDateIndex(int chekcinDateIndex) {
		this.chekcinDateIndex = chekcinDateIndex;
	}
	
	public int getPriceIndex() {
		return priceIndex;
	}
	
	public void setPriceIndex(int priceIndex) {
		this.priceIndex = priceIndex;
	}
	
	public int getHotelNameIndex() {
		return hotelNameIndex;
	}
	
	public void setHotelNameIndex(int hotelNameIndex) {
		this.hotelNameIndex = hotelNameIndex;
	}
	
	public int getCurrencyIndex() {
		return currencyIndex;
	}
	
	public void setCurrencyIndex(int currencyIndex) {
		this.currencyIndex = currencyIndex;
	}
	
	public String getSubjectHotel() {
		return subjectHotel;
	}

	public void setSubjectHotel(String subjectHotel) {
		this.subjectHotel = subjectHotel;
	}
	
	public boolean isSkipHeadline() {
		return skipHeadline;
	}

	public void setSkipHeadline(boolean skipHeadline) {
		this.skipHeadline = skipHeadline;
	}

	public String getInput() {
		return input;
	}

	public void setInput(String input) {
		this.input = input;
	}
	
	

	public String getProviderName() {
		return providerName;
	}

	public void setProviderName(String providerName) {
		this.providerName = providerName;
	}

	@Override
	public Date parseDateField(Cell cell, String dateFormat) {

		DateFormat format = new SimpleDateFormat(dateFormat);

		Date date = null;
		String dateStr = null;
		int cellType = cell.getCellType();
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
				try {
					date = format.parse(dateStr);
				} catch (ParseException e) {
					e.printStackTrace();
				}
			}

		} else if (cellType == 0) {
			date = cell.getDateCellValue();
		}
		
		return date;
	}

	@Override
	public String parseStringField(Cell cell) {
		return cell.getStringCellValue();
	}

	@Override
	public Double parseNumericField(Cell cell) {
		return cell.getNumericCellValue();
	}

	
}
