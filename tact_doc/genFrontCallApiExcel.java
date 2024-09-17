import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.util.Scanner;


import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
//
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.io.IOUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
//

public class genFrontCallApiExcel {
    
	private static List<String> fileList = new ArrayList<>();    	
	private static List<String[]> fieldList = new ArrayList<>(); 
	private static String prefixWordName = "";
	private static int chapNo = 1;  // 章節編號
	
	private static String chapTitle = "";
	private static String requestMapping = "";
	
	private static String sectionTitle = "";
	private static String apiDesc = "";
	private static String apiMethod = "";
	private static String apiURL = "";
	private static String paramName = "";
	private static String paramObject = "";
	private static String mappingStr = "";
	
	
	private static final int FILENAME = 0;
	private static final int METHOD = 1;
	private static final int URL = 2;
	private static final int MAPPING = 3;
		    
    private static final String[] titles = {
            "FILENAME", "METHOD", "URL"};
	
    public static void appendNewPage(String wordFilename) {
    	
    }
	
	
    public static void getFileList(String path) throws Exception  {
    	
    	//System.out.println("=== scan path=" + path);
        Scanner s1 = new Scanner(path);
        
        String folderPath = s1.next();
        File folder = new File(folderPath);
        
        if (folder.isDirectory()) {
           File[] listOfFiles = folder.listFiles();
           
           for (File file : listOfFiles) {
              if(file.isDirectory()) {
            	  getFileList(file.getCanonicalPath().toString());            	  
              } else {

            	  fileList.add(file.getCanonicalPath().toString());
            	  //System.out.println(file.getCanonicalPath().toString());
              }	
            	
           } 
        }     	
    }
    
    public static void clearVarAll() {
    	sectionTitle = "";
    	apiDesc = "";
    	apiMethod = "";
    	apiURL = "";
    	paramName = "";
    	paramObject = "";    	
    	mappingStr = "";
    }
    
    public static void parseFile(String fileName) throws Exception  {
    	FileReader fr = new FileReader(fileName);
    	BufferedReader br = new BufferedReader(fr);
    	int sectionNo = 1;
    	String shortFileName = fileName;    	

    	int lastSlashPos = fileName.lastIndexOf("\\");
    	if (lastSlashPos >= 0 && fileName.length() > lastSlashPos + 1) {
    		shortFileName = fileName.substring(lastSlashPos + 1);
    	}
    	System.out.println("shortFileName=" + shortFileName);
    	while (br.ready()) {
    		String line = br.readLine().trim();
    		int pos = 0;
    		
    		//Axios.get(process.env.REACT_APP_API_SERVER_EXP + 'dangerGoodsTable/getList', {
    		int axiosPos = line.indexOf("Axios");
    		if (axiosPos >= 0) {
    			if (line.startsWith("import")){
    				continue;
    			}	
    			int bracketPos = line.indexOf("(");
    			if (bracketPos < 0){
    				System.out.println("error!");
    				continue;
    			}
    			
    			int markPos = line.indexOf("'");
    			if (markPos < 0){
    				System.out.println("error!");
    				continue;
    			}
    			if (line.length() <= markPos + 1) {
    				System.out.println("error!");
    				continue;    				
    			}
    			
    			System.out.println("line=" + line);
    			
    			// method
    			String method = line.substring(axiosPos + "Axios.".length(), bracketPos);
    			    			
    			// url
    			String url = line.substring(markPos + 1);
    			pos = url.indexOf("'");
    			if (pos >= 0) {
    				url = url.substring(0, pos);
    			}
    			System.out.println("method=" + method + ",url=" + url);
    			
				String saFieldInfo[]= new String[3];
				saFieldInfo[FILENAME] = shortFileName;
				saFieldInfo[METHOD] = method;
				saFieldInfo[URL] = url;
				            		
				fieldList.add(saFieldInfo);									
				//\clearVarAll();    			
    		}
		

    	}
    	fr.close();    	
    }
    
 
    public static void main(String[] args) throws Exception {
        Workbook wb;
        int excelRowNum = 1, excelSheet2RowNum = 0;
    	
    	
    	//prefixWordName = "I626-API-005_CH2.";
        getFileList("C:\\我的資料(d)\\GIT_Source_Front_Doc\\EXP");
                        
        // excel
        wb = new HSSFWorkbook();
        Map<String, CellStyle> styles = createStyles(wb);
        Sheet sheet = wb.createSheet("Sheet1");
        Sheet sheet2 = wb.createSheet("Sheet2");
        sheet.setAutobreaks(true);
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < titles.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(titles[i]);
            cell.setCellStyle(styles.get("header"));
        }
        
        ///

System.out.println("start");
        for (String filename: fileList) {
        	System.out.println(filename);
        	parseFile(filename);
			String saWordFilelist[]= new String[fieldList.size()];
        	for (int i = 0; i < fieldList.size(); i++) {
            	String outputFilename = "output_api_word\\" + prefixWordName + chapNo + "." + (i + 1) + ".docx";
            	saWordFilelist[i] = outputFilename;
        		
        		// genExcel
                Row row;
                Cell cell;
                row = sheet.createRow(excelRowNum++);
                for (int j = 0; j < fieldList.get(i).length; j++) {
                	cell = row.createCell(j);
                	cell.setCellValue(fieldList.get(i)[j]);
                	cell.setCellStyle(styles.get("cell_data"));                 	
                }
                ////
                                

        	}        	
        	
        	fieldList.clear();
        	chapTitle = "";
        	requestMapping = "";        	
        	chapNo++;
        }
        
        // excel file
        String outExcelName = "output_front_callapi\\summary.xls";
        FileOutputStream out = new FileOutputStream(outExcelName);        
        wb.write(out);
        out.close();
        wb.close();
        
        
        System.out.println("執行完畢!");
    	
    	
    }
    
    /**
     * create a library of cell styles
     */
    private static Map<String, CellStyle> createStyles(Workbook wb){
        Map<String, CellStyle> styles = new HashMap<>();
        DataFormat df = wb.createDataFormat();

        CellStyle style;
        Font headerFont = wb.createFont();
        headerFont.setBold(true);
        headerFont.setFontName("標楷體");
        
        style = createBorderedStyle(wb);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(headerFont);
        styles.put("header", style);

        Font fontData = wb.createFont();
        fontData.setFontName("標楷體");
        //fontData.setBold(false);

        style = createBorderedStyle(wb);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(fontData);
        styles.put("cell_data", style);
        
        
        style = createBorderedStyle(wb);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(headerFont);
        style.setDataFormat(df.getFormat("d-mmm"));
        styles.put("header_date", style);

        Font font1 = wb.createFont();
        font1.setBold(true);
        style = createBorderedStyle(wb);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setFont(font1);
        styles.put("cell_b", style);

        style = createBorderedStyle(wb);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFont(font1);
        styles.put("cell_b_centered", style);

        style = createBorderedStyle(wb);
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setFont(font1);
        style.setDataFormat(df.getFormat("d-mmm"));
        styles.put("cell_b_date", style);

        style = createBorderedStyle(wb);
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setFont(font1);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setDataFormat(df.getFormat("d-mmm"));
        styles.put("cell_g", style);

        Font font2 = wb.createFont();
        font2.setColor(IndexedColors.BLUE.getIndex());
        font2.setBold(true);
        style = createBorderedStyle(wb);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setFont(font2);
        styles.put("cell_bb", style);

        style = createBorderedStyle(wb);
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setFont(font1);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setDataFormat(df.getFormat("d-mmm"));
        styles.put("cell_bg", style);

        Font font3 = wb.createFont();
        font3.setFontHeightInPoints((short)14);
        font3.setColor(IndexedColors.DARK_BLUE.getIndex());
        font3.setBold(true);
        style = createBorderedStyle(wb);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setFont(font3);
        style.setWrapText(true);
        styles.put("cell_h", style);

        style = createBorderedStyle(wb);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setWrapText(true);
        styles.put("cell_normal", style);

        style = createBorderedStyle(wb);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setWrapText(true);
        styles.put("cell_normal_centered", style);

        style = createBorderedStyle(wb);
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setWrapText(true);
        style.setDataFormat(df.getFormat("d-mmm"));
        styles.put("cell_normal_date", style);

        style = createBorderedStyle(wb);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setIndention((short)1);
        style.setWrapText(true);
        styles.put("cell_indented", style);

        style = createBorderedStyle(wb);
        style.setFillForegroundColor(IndexedColors.BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styles.put("cell_blue", style);
        
        return styles;
    }

    private static CellStyle createBorderedStyle(Workbook wb){
        BorderStyle thin = BorderStyle.THIN;
        short black = IndexedColors.BLACK.getIndex();

        CellStyle style = wb.createCellStyle();
        style.setBorderRight(thin);
        style.setRightBorderColor(black);
        style.setBorderBottom(thin);
        style.setBottomBorderColor(black);
        style.setBorderLeft(thin);
        style.setLeftBorderColor(black);
        style.setBorderTop(thin);
        style.setTopBorderColor(black);
        return style;
    }
}
