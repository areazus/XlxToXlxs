package areaz.us.converter.xlsToXlsx;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Timestamp;
import java.util.logging.FileHandler;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;

import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import jxl.Sheet;
import jxl.Workbook;


public class XlsConverter {
	
	//Logger Initialization
	private static final Logger LOGGER = initLogger();
	public static Logger initLogger(){
		Logger logger = Logger.getLogger(XlsConverter.class.getName());
		Timestamp time = new Timestamp(System.currentTimeMillis());
		try {
			FileHandler fh = new FileHandler(XlsConverter.class.getName()+time.toString().replaceAll("[^\\d.]", "")+".log");
			logger.setUseParentHandlers(false);
			logger.addHandler(fh);
			SimpleFormatter formatter = new SimpleFormatter();
            fh.setFormatter(formatter);
		} catch (SecurityException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return logger;
	}
	//End of Logger initialization
	
	private final File file;
	
	public XlsConverter(String fileName){
		this(new File(fileName));
	}
	
	public XlsConverter(File mfile){
		this.file = mfile;
		LOGGER.info("Trying to convert "+file.getAbsolutePath());
		if(file !=null && file.exists() && file.getName().endsWith(".xls")){
			convertFile();
		}else{
			LOGGER.severe(file.getAbsolutePath() + " doesn't exist OR not xls file!");
		}
	}
	
	private void convertFile(){
		try{
			//ReadFile
			Workbook readWorkBook = Workbook.getWorkbook(file);
			int sheets = readWorkBook.getNumberOfSheets();
			LOGGER.info("Found "+sheets+" sheets");
			//OutputFile
			String outputFilePath = file.getAbsolutePath()+"x";
			FileOutputStream outputFile = new FileOutputStream(file.getAbsolutePath()+"x");
			XSSFWorkbook outWorkBook = new XSSFWorkbook();

			for(int i=0; i<sheets; i++){
				Sheet readSheet = readWorkBook.getSheet(i);
				int rows = readSheet.getRows();
				int colums = readSheet.getColumns();
				LOGGER.info("Found "+rows+" rows and "+colums+" columns in sheet "+(i+1));
				XSSFSheet outputSheet = outWorkBook.createSheet(readSheet.getName());
				for(int j=0; j<rows; j++){
					XSSFRow row = outputSheet.createRow(j);
					for(int k=0; k<colums; k++){
						String cellContent = readSheet.getCell(k,j).getContents();
						XSSFCell cell = row.createCell(k);
						cell.setCellValue(cellContent);
					}
				}
				LOGGER.info("Done with "+(i+1)+"/"+sheets+" sheets");
			}
			outWorkBook.write(outputFile);
			outputFile.flush();
			outputFile.close();
			LOGGER.info("Completed task: Output saved to "+outputFilePath);
		}catch (Exception e){
			LOGGER.log(Level.SEVERE, e.getMessage(), e);
		}
	}
	
	public static File getXlXSFile(String filePath){
		if(filePath !=null){
			if(filePath.endsWith(".xlsx")){
				LOGGER.warning("File already seems to be xlsx file, skipping conversion");
				return new File(filePath);
			}else if(filePath.endsWith(".xls")){
				String outputFilePath = filePath+"x";
				new XlsConverter(filePath);
				File toReturn = new File(outputFilePath);
				if(toReturn.exists()){
					LOGGER.info("Conversion success, will return xlsx file");
					return toReturn;
				}else{
					LOGGER.severe("Conversion failed, will return null");
				}
			}
		}
		LOGGER.severe("FilePath is null or not xls file, or conversion failed, will return null");
		return null;
	}

	public static void main(String[] agr){
		new XlsConverter("C:\\Users\\ahmed\\Downloads\\download.xls");
	}
}
