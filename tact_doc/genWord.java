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

public class genWord {
    
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
	
	
	private static final int CHAP_TITLE = 0;
	private static final int SECTION_TITLE = 1;
	private static final int API_DESC = 2;
	private static final int API_METHOD = 3;
	private static final int API_URL = 4;
	private static final int PARAM_NAME = 5;
	private static final int PARAM_OBJECT = 6;
	private static final int MAPPING = 7;
		    
    private static final String[] titles = {
            "CHAP_TITLE", "SECTION_TITLE", "API_DESC", "API_METHOD", "API_URL", "PARAM_NAME", "PARAM_OBJECT"};
	
    public static void appendNewPage(String wordFilename) {
    	
    }
	
	/**
	 * 合併docx檔案
	 * @param srcDocxs 需要合併的目標docx檔案
	 * @param destDocx 合併後的docx輸出檔案
	 */
	public static void mergeDoc(String[] srcDocxs,String destDocx){
		
		OutputStream dest = null;
		List<OPCPackage> opcpList = new ArrayList<OPCPackage>();
		int length = null == srcDocxs ? 0 : srcDocxs.length;
		/**
		 * 迴圈獲取每個docx檔案的OPCPackage物件
		 */
		for (int i = 0; i < length; i++) {
			String doc = srcDocxs[i];
			OPCPackage srcPackage =  null;
			try {
				srcPackage = OPCPackage.open(doc);
			} catch (Exception e) {
				e.printStackTrace();
			}
			if(null != srcPackage){
				opcpList.add(srcPackage);
			}
		}
		
		int opcpSize = opcpList.size();
		//獲取的OPCPackage物件大於0時，執行合併操作
		if(opcpSize > 0){
			try {
				dest = new FileOutputStream(destDocx);
				XWPFDocument src1Document = new XWPFDocument(opcpList.get(0));
								
				CTBody src1Body = src1Document.getDocument().getBody();
				//OPCPackage大於1的部分執行合併操作
				if(opcpSize > 1){
					for (int i = 1; i < opcpSize; i++) {
						OPCPackage src2Package = opcpList.get(i);
						XWPFDocument src2Document = new XWPFDocument(src2Package);
						CTBody src2Body = src2Document.getDocument().getBody();
						appendBody(src1Body, src2Body);
					}
				}
				//將合併的文件寫入目標檔案中
				src1Document.write(dest);
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			} catch (Exception e) {
				e.printStackTrace();
			}finally{
                //註釋掉以下部分，去除影響目標檔案srcDocxs。
				/*for (OPCPackage opcPackage : opcpList) {
					if(null != opcPackage){
						try {
							opcPackage.close();
						} catch (IOException e) {
							e.printStackTrace();
						}
					}
				}*/
				//關閉流
				IOUtils.closeQuietly(dest);
			}
		}
		
		
	}
	
	/**
	 * 合併文件內容
	 * @param src 目標文件
	 * @param append 要合併的文件
	 * @throws Exception
	 */
	private static void appendBody(CTBody src, CTBody append) throws Exception {
		XmlOptions optionsOuter = new XmlOptions();
		optionsOuter.setSaveOuter();
		String appendString = append.xmlText(optionsOuter);
		String srcString = src.xmlText();
		String prefix = srcString.substring(0, srcString.indexOf(">") + 1);
		String mainPart = srcString.substring(srcString.indexOf(">") + 1,
				srcString.lastIndexOf("<"));
		String sufix = srcString.substring(srcString.lastIndexOf("<"));
		String addPart = appendString.substring(appendString.indexOf(">") + 1,
				appendString.lastIndexOf("<"));
		CTBody makeBody = CTBody.Factory.parse(prefix + mainPart + addPart
				+ sufix);
		src.set(makeBody);
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
            	  // 關貿開發的程式, 跳過
            	  if (file.getCanonicalPath().toString().contains("exp\\CityNationAirlineController.java")) {
            		  continue;
            	  }

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

    	while (br.ready()) {
    		String line = br.readLine().trim();
    		int pos = 0;
    		
    		// @Api(description = "")    		
    		if ((line.indexOf("@Api(") >= 0 || line.indexOf("@Api ") >= 0) && line.indexOf("description") >= 0) {
    			pos = line.indexOf("=");
    			String tmpLine = line.substring(pos);
    			tmpLine = tmpLine.replace(" ", "");
    			tmpLine = tmpLine.replace("=", "");
    			tmpLine = tmpLine.replace("(", "");
    			tmpLine = tmpLine.replace(")", "");
    			tmpLine = tmpLine.replace("\"", "");
    			chapTitle = "2." + chapNo + tmpLine;
    		}
    		
    		// @RequestMapping("")
    		if ((pos = line.indexOf("@RequestMapping")) >= 0) {
    			String tmpLine = line.substring(pos + "@RequestMapping".length());    			
    			tmpLine = tmpLine.replace(" ", "");
    			tmpLine = tmpLine.replace("(", "");
    			tmpLine = tmpLine.replace(")", "");
    			tmpLine = tmpLine.replace("\"", "");
    			requestMapping = tmpLine;    			
    		}
    		
    		// @ApiOperation(value = "查詢異常代碼資料")
    		if ((pos = line.indexOf("@ApiOperation")) >= 0) {
    			String tmpLine = line.substring(pos + "@ApiOperation".length()).trim();
    			tmpLine = tmpLine.replace(" ", "");
    			//tmpLine = tmpLine.replace("(", "");
    			//tmpLine = tmpLine.replace(")", "");
    			//tmpLine = tmpLine.replace("\"", "");
    			String tmpArr[] = tmpLine.split("=");
    			if (tmpArr.length >= 2) {
    				//if (tmpArr[0].equalsIgnoreCase("value")) {
    				if (tmpArr[0].contains("value")){
    					String tmpName = tmpArr[1];
    					tmpName = tmpName.replace("\")", "");
    					tmpName = tmpName.replace("\"", "");
    					sectionTitle =  "2." + chapNo + "." + sectionNo + tmpName;    
    					apiDesc = tmpName;
    				}    				
    			}
    		}
    				
    		// @GetMapping("")  @PostMapping("")  @PutMapping("") @DeleteMapping("")
    		if (line.indexOf("@GetMapping") >= 0 || line.indexOf("@PostMapping") >= 0 || line.indexOf("@PutMapping") >= 0 || line.indexOf("@DeleteMapping") >= 0 ) {
    			String tmpLine = "";
    			if ((pos = line.indexOf("@GetMapping")) >= 0) {
        			tmpLine = line.substring(pos + "@GetMapping".length());   
        			apiMethod = "GET";
    			} else if ((pos = line.indexOf("@PostMapping")) >= 0) {
        			tmpLine = line.substring(pos + "@PostMapping".length());
        			apiMethod = "POST";
    			} else if ((pos = line.indexOf("@PutMapping")) >= 0) {
        			tmpLine = line.substring(pos + "@PutMapping".length());
        			apiMethod = "PUT";
    			} else if ((pos = line.indexOf("@DeleteMapping")) >= 0) {
        			tmpLine = line.substring(pos + "@DeleteMapping".length());
        			apiMethod = "DELETE";
    			}
    			tmpLine = tmpLine.replace(" ", "");
    			tmpLine = tmpLine.replace("(", "");
    			tmpLine = tmpLine.replace(")", "");
    			tmpLine = tmpLine.replace("\"", "");
    			if (tmpLine.startsWith("/")) {
        			apiURL = requestMapping + tmpLine;     				
    			} else if (tmpLine.trim().equals("")){
    				apiURL = requestMapping;
    			} else {	
        			apiURL = requestMapping + "/" + tmpLine;     				
    			}
    			mappingStr = tmpLine;
    			//System.out.println("apiMethod=" + apiMethod);
    			//System.out.println("apiURL=[" + apiURL + "]");    			
    		}

    		// check error
    		if (line.indexOf("public") >= 0 && line.indexOf("(") >= 0 && line.indexOf(")") < 0 && line.indexOf(";") < 0) {
    			System.out.println("check error ==> line=" + line);
    		}
    		
    		
    		// public ResponseEntity updateAbnormalTb(@Valid @RequestBody UpdateAbnormalTbRequest request)
    		if (line.indexOf("public") >= 0 && line.indexOf("(") >= 0 && line.indexOf(")") >= 0 && line.indexOf(";") < 0) {
    			pos = line.indexOf("(");
    			String tmpLine = line.substring(pos);   
    			tmpLine = tmpLine.replace("(", "");
    			tmpLine = tmpLine.replace(")", "");
    			tmpLine = tmpLine.replace("\"", "");
    			String tmpArr[] = tmpLine.split(" ");
    			for (int i = 0; i < tmpArr.length; i++) {
    				if (tmpArr[i].startsWith("@")) {
    					continue;
    				} else {
    					paramName = tmpArr[i];
    					paramObject = tmpArr[i];
    					break;
    				}
    				
    			}
    			////// 資料已抓完
				String saFieldInfo[]= new String[8];
				if (sectionNo > 1)
					saFieldInfo[CHAP_TITLE] = "";
				else
					saFieldInfo[CHAP_TITLE] = chapTitle;
				saFieldInfo[SECTION_TITLE] = sectionTitle;
				saFieldInfo[API_DESC] = apiDesc;
				saFieldInfo[API_METHOD] = apiMethod;
				saFieldInfo[API_URL] = apiURL;
				saFieldInfo[PARAM_NAME] = paramName;
				saFieldInfo[PARAM_OBJECT] = paramObject;
				saFieldInfo[MAPPING] = mappingStr;				
				fieldList.add(saFieldInfo);									
				clearVarAll();
				sectionNo++;	  
    		}
    	}
    	fr.close();    	
    }
    
    public static void genWord(String[] saData, String outputFilename) throws Exception {
    	FileOutputStream fos = null;    	
    	String templateFilename = "";
    	String firstPage = "";
    	String logicDesc = "";
    	
    	//System.out.println("saData[API_DESC]=" + saData[API_DESC]);
    	if (!saData[CHAP_TITLE].equals("")){
    		firstPage = "_firstpage";
    	}
    	if (saData[API_METHOD].equals("GET") && saData[PARAM_NAME].equals("")) {
    		templateFilename = "template_get_no_param" + firstPage + ".docx";
    		if (saData[MAPPING].equals("")) {
    			logicDesc = "以提供的參數，針對Table進行資料查詢，回傳查詢結果。";
    		} else {
    			logicDesc = "以提供的參數，進行資料查詢，查詢SQL如下:";    			
    		}
    	} else if (saData[API_METHOD].equals("GET")) {
        	templateFilename = "template_get" + firstPage + ".docx";   		
    		if (saData[MAPPING].equals("")) {
    			logicDesc = "以提供的參數，針對Table進行資料查詢，回傳查詢結果。";
    		} else {
    			logicDesc = "以提供的參數，進行資料查詢，查詢SQL如下:";    			
    		}
    	} else if (saData[API_METHOD].equals("POST")){
    		templateFilename = "template_post" + firstPage + ".docx";   		
    	} else if (saData[API_METHOD].equals("PUT")){
    		templateFilename = "template_put" + firstPage + ".docx";   		
    	} else if (saData[API_METHOD].equals("DELETE")){
    		templateFilename = "template_delete" + firstPage + ".docx";    		    		
    	}
    	XWPFDocument document = new XWPFDocument(OPCPackage.open("template\\" + templateFilename));
    	for (XWPFParagraph paragraph : document.getParagraphs()) {    	
    	    for (XWPFRun run : paragraph.getRuns()) {
    	        String text = run.text();
    	        if (text != null) {
	    	        text = text.replace("ChapTitle", saData[CHAP_TITLE]);
	    	        text = text.replace("SectionTitle", saData[SECTION_TITLE]);
	    	        text = text.replace("APIDesc", saData[API_DESC]);
	    	        run.setText(text,0);
	    	        //System.out.println(text);
    	        }    	        
    	    }    	    
    	}
    	for (XWPFTable table : document.getTables()) {
    	       List<XWPFTableRow> rowList = table.getRows();
    	       for (int i = 0; i < rowList.size(); i++) {
    	           List<XWPFTableCell> cellList = rowList.get(i).getTableCells();
    	           for (int j = 0; j < cellList.size(); j++) {
    	               XWPFParagraph cellParagraph = cellList.get(j).getParagraphArray(0);
    	       	       for (XWPFRun run : cellParagraph.getRuns()) {
        	               String text = run.text();
        	               if (text != null) {
        		    	       text = text.replace("APIURL", saData[API_URL]);  
        		    	       text = text.replace("APIMethod", saData[API_METHOD]);
        		    	       text = text.replace("PARAM_NAME", saData[PARAM_NAME]);
        		    	       text = text.replace("PARAM_OBJECT", saData[PARAM_OBJECT]);
        		    	       text = text.replace("LogicDesc", logicDesc);        		    	       
        		    	       run.setText(text,0);             		    	       
        	               }    	       	    	   
    	       	       }
    	           }    	    	   
    	       }	               	       
    	}
    	    	
    	fos = new FileOutputStream(outputFilename);
    	document.write(fos);  
    	fos.flush();
    	fos.close();
    	//document.close();
    }
    
    
//    https://poi.apache.org/casestudies.html    
 // https://www.itread01.com/content/1544970672.html
 // https://www.tutorialspoint.com/apache_poi_word/apache_poi_word_tables.htm  
//    https://www.itread01.com/content/1544970672.html    
    public static void main(String[] args) throws Exception {
        Workbook wb;
        int excelRowNum = 1, excelSheet2RowNum = 0;
    	
    	//genWord("GET", chapTitle, sectionTitle, APIDesc);
    	
    	/*String[] srcDocxs = {"1.docx","2.docx","3.docx"};
		String destDocx = "0.docx";
		mergeDoc(srcDocxs, destDocx);*/
    	
        // edi
    	//prefixWordName = "I626-API-006-002_CH2.";
        //getFileList("C:\\我的資料(d)\\GIT_Source_Back_EDI\\src\\main\\java\\com\\tradevan\\twms\\web\\controller\\edi02");
        
        // dis
    	prefixWordName = "I626-API-005_CH2.";
        getFileList("C:\\我的資料(d)\\GIT_Source_Back_DIS\\src\\main\\java\\com\\tradevan\\twms\\web\\controller\\dis");
        
        // exp
    	//prefixWordName = "I626-API-004_CH2.";
        //getFileList("C:\\我的資料(d)\\GIT_Source_Back_EXP\\src\\main\\java\\com\\tradevan\\twms\\web\\controller\\exp");
        
        // text file
        FileWriter fw = new FileWriter("output_api_summary\\" + prefixWordName + "summary.txt");        
        
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
            fw.write(chapTitle + "\n");
        	for (int i = 0; i < fieldList.size(); i++) {
            	String outputFilename = "output_api_word\\" + prefixWordName + chapNo + "." + (i + 1) + ".docx";
            	saWordFilelist[i] = outputFilename;
        		genWord(fieldList.get(i), outputFilename);  
        		
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
                                
                String strTextOut = fieldList.get(i)[API_DESC] + "," + fieldList.get(i)[API_METHOD] + "," + fieldList.get(i)[API_URL] + "," + fieldList.get(i)[PARAM_NAME];
                fw.write(strTextOut + "\n");

        	}
        	String mergeDesFilename = "output_api_word_merge\\" + prefixWordName + chapNo + ".docx";
        	mergeDoc(saWordFilelist, mergeDesFilename);
        	
        	// excel file
            Row row2;
            Cell cell2;
            row2 = sheet2.createRow(excelSheet2RowNum++);
        	cell2 = row2.createCell(0);
        	cell2.setCellValue(chapTitle);
        	cell2.setCellStyle(styles.get("cell_data"));                 	
            
        	// text file
        	fw.write("=================================\n");        	        	
        	
        	fieldList.clear();
        	chapTitle = "";
        	requestMapping = "";        	
        	chapNo++;
        }
        
        // excel file
        String outExcelName = "output_api_summary\\" + prefixWordName+ "summary.xls";
        FileOutputStream out = new FileOutputStream(outExcelName);        
        wb.write(out);
        out.close();
        wb.close();
        
        // text file
        fw.flush();
        fw.close();
        
        System.out.println("執行完畢!");
    	
    	
    }
    
/*
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
    } */

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
