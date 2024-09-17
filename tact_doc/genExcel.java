import java.io.FileOutputStream;
import java.io.FileReader;
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



import org.apache.poi.hssf.usermodel.HSSFWorkbook;
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




public class genExcel {

    private static final String[] titles = {
            "欄位名稱", "資料型別", "必要", "備註"};
    
	private static List<String> fileList = new ArrayList<>();    	
	private static List<String[]> fieldList = new ArrayList<>(); 
	private static String className = "";
	private static String varChiName = "";
	private static String varRequired = "";
	private static String varEngName = "";
	private static String varType = "";
    
    
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
    	varChiName = "";
    	varRequired = "";
    	varEngName = "";
    	varType = "";    	
    }
    
    public static void parseFile(String fileName) throws Exception  {
    	FileReader fr = new FileReader(fileName);
    	BufferedReader br = new BufferedReader(fr);

    	while (br.ready()) {
    		String line = br.readLine();
    		System.out.println(line);
    		int pos = 0;
    		
    		// class
    		if ((pos = line.indexOf(" class ")) >= 0) {
    			String tmpLine = line.substring(pos + " class ".length());
    			String tmpArray[] = tmpLine.split(" ");
    			if (tmpArray.length > 0) {
    				className = tmpArray[0];
    			}
    			if (line.indexOf(" extends ") >= 0) {
    				if (line.indexOf(" GetPageRequest") >= 0) {
    					////// 資料已抓完
    					String saFieldInfo1[]= new String[4];
    					String saFieldInfo2[]= new String[4];
    					saFieldInfo1[0] = "page";
    					saFieldInfo1[1] = "Integer";
    					saFieldInfo1[2] = "是";
    					saFieldInfo1[3] = "查詢的資料頁數";
    					fieldList.add(saFieldInfo1);
    					
    					saFieldInfo2[0] = "pageSize";
    					saFieldInfo2[1] = "Integer";
    					saFieldInfo2[2] = "是";
    					saFieldInfo2[3] = "分頁筆數";
    					fieldList.add(saFieldInfo2);
    				} else {
    					System.out.println("michelle extends ==>" + line);
    				}
    			}
    		}
    		    		
    		// 處理變數中文意義及是否必要欄
    		if ((pos = line.indexOf("@ApiModelProperty")) >= 0) {
    			String tmpLine = line.substring(pos + "@ApiModelProperty".length());
    			tmpLine = tmpLine.replace(" ", "");
    			tmpLine = tmpLine.replace("(", "");
    			tmpLine = tmpLine.replace(")", "");
    			tmpLine = tmpLine.replace("\"", "");
    			
    			String tmpArray[] = tmpLine.split(",");
    			for (String element: tmpArray) {
    				String tmpArr[] = element.split("=");
    				if (tmpArr.length >= 2) {
    					varRequired = "否";
    					if (tmpArr[0].equals("value")) {
    						varChiName = tmpArr[1];
    					}
    					if (tmpArr[0].equals("required")) {
    						if (tmpArr[1].equalsIgnoreCase("true")) {
    							varRequired = "是";    				
    						}	
    					}
    				}
    			}
    		}
    		
		    //private String abnormal;			
    		// 處理變數名稱及資料型別
    		if ((pos = line.indexOf(";")) >= 0) {
    			String tmpLine = line.substring(0, pos);
    			String tmpArr[] = tmpLine.split(" ");
    			for (int i = 0; i < tmpArr.length; i++) {
    				if (tmpArr[i].equals("private") || tmpArr[i].equals("public")) {
    					if (tmpArr.length >= i + 2 + 1) {
        					varType = tmpArr[i + 1];
        					varEngName = tmpArr[i + 2];
        					
        					////// 資料已抓完
        					String saFieldInfo[]= new String[4];
        					saFieldInfo[0] = varEngName;
        					saFieldInfo[1] = varType;
        					saFieldInfo[2] = varRequired;
        					saFieldInfo[3] = varChiName;
        					fieldList.add(saFieldInfo);
        					
        					clearVarAll();
        					        					
    					}	
    					break;
    				}	    				
    			}
    			System.out.println("tmpLine=" + tmpLine);
    			System.out.println("varType=" + varType);
    			System.out.println("varEngName=" + varEngName);    			
    		}
    		
    		
    	}
    	fr.close();
    	for (String[] varArray: fieldList) {
    		System.out.println("varArray=[" + varArray[0] + "][" + varArray[1] + "][" + varArray[2] + "][" + varArray[3] +"]");
    	}
    	
    }
    

    public static void main(String[] args) throws Exception {
        Workbook wb;
        
        getFileList("C:\\我的資料(d)\\GIT_Source_Back_EDI\\src\\main\\java\\com\\tradevan\\twms\\web\\model");
        
        int count = 0;
        for (String filename: fileList) {
        	System.out.println(filename);
        	parseFile(filename);
        	if (count++ > 20)
        		break;
            wb = new HSSFWorkbook();
            Map<String, CellStyle> styles = createStyles(wb);
            Sheet sheet = wb.createSheet("Sheet1");
            sheet.setAutobreaks(true);
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < titles.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(titles[i]);
                cell.setCellStyle(styles.get("header"));
            }
            Row row;
            Cell cell;
            int rownum = 1;
        	for (String[] varArray: fieldList) {
                row = sheet.createRow(rownum++);
                for (int j = 0; j < 4; j++) {
                    cell = row.createCell(j);
                    cell.setCellValue(varArray[j]);
                    cell.setCellStyle(styles.get("cell_data"));                    
                }    
        	}
            sheet.setColumnWidth(0, 256*16);
            sheet.setColumnWidth(1, 256*10);  
            sheet.setColumnWidth(2, 256*6);
            sheet.setColumnWidth(3, 256*34);
            
            String file = className  + ".xls";
            FileOutputStream out = new FileOutputStream(file);
            wb.write(out);
            out.close();
            wb.close();
            fieldList.clear();
        }
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
