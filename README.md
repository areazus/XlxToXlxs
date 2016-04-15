A very simple java converter to convert xls to xlsx files. Helpful when processing datasheets ust be done with POI, and org.apache.poi.hssf.OldExcelFormatException error occurs. Converts the cell content as String and save into a new excel file. All formatting is lost.

POI message for old excel files: "The supplied spreadsheet seems to be Excel 5.0/7.0 (BIFF5) format. POI only supports BIFF8 format (from Excel versions 97/2000/XP/2003)"

Usage:
	new areaz.us.converter.xlsToXlsx.XlsConverter(Path To File);
	
Feel free to Fork or suggest changes.