package parsers;

import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;

public interface CSVParser {

	Date parseDateField(Cell cell, String dateFormat);
	String parseStringField(Cell cell);
	Double parseNumericField(Cell cell);
}
