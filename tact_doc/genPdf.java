import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.*;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import java.io.*;
import java.util.*;

public class genPdf {
    
	/**
     * @param map 需要置換的欄位
     * @param sourceFile  來源文件路徑
     * @param targetFile  目的文件路徑
     * @throws IOException
     */
    public static void genPdf(HashMap map, String sourceFile, String targetFile) throws IOException {
        File templateFile = new File(sourceFile);
        fillParam(map, FileUtils.readFileToByteArray(templateFile), targetFile);
    }

    /**
     * 使用map中的參數填充pdf，map中的key和pdf檔中的field對應
     */
    public static void fillParam(Map<String, String> fieldValueMap, byte[] file, String contractFileName) {
        FileOutputStream fos = null;
        try {
        
            fos = new FileOutputStream(contractFileName);
            PdfReader reader = null;
            PdfStamper stamper = null;
            BaseFont base = null;
            try {
                reader = new PdfReader(file);
                stamper = new PdfStamper(reader, fos);
                stamper.setFormFlattening(true);
                base = BaseFont.createFont("kaiu.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                
                AcroFields acroFields = stamper.getAcroFields();
                for (String key : acroFields.getFields().keySet()) {
                    acroFields.setFieldProperty(key, "textfont", base, null);
                    acroFields.setFieldProperty(key, "textsize", new Float(10), null);   //字體大小
                }
                if (fieldValueMap != null) {
                    for (String fieldName : fieldValueMap.keySet()) {
                        if (StringUtils.isNotBlank(fieldValueMap.get(fieldName))) {
                            //取得map中key對應的Value是否為On, 若是則勾選複選框
                            if (fieldValueMap.get(fieldName).equals("On") || fieldValueMap.get(fieldName) == "On") {
                                acroFields.setField(fieldName, fieldValueMap.get(fieldName),true);
                            }else {
                                acroFields.setField(fieldName, fieldValueMap.get(fieldName));
                            }
                        }
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
            } finally {
                if (stamper != null) {
                    try {
                        stamper.close();
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
                if (reader != null) {
                    reader.close();
                }
            }

        } catch (Exception e) {
            System.out.println("填充參數異常");
            e.printStackTrace();
        } finally {
            IOUtils.closeQuietly(fos);
        }
    }


    /**
     * 取得pdf中的fieldNames
     */
    public static Set<String> getTemplateFileFieldNames(String pdfFileName) {
        Set<String> fieldNames = new TreeSet<String>();
        PdfReader reader = null;
        try {
            reader = new PdfReader(pdfFileName);
            Set<String> keys = reader.getAcroFields().getFields().keySet();
            for (String key : keys) {
                int lastIndexOf = key.lastIndexOf(".");
                int lastIndexOf2 = key.lastIndexOf("[");
                fieldNames.add(key.substring(lastIndexOf != -1 ? lastIndexOf + 1 : 0, lastIndexOf2 != -1 ? lastIndexOf2 : key.length()));
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (reader != null) {
                reader.close();
            }
        }

        return fieldNames;
    }


    /**
     * 讀取文件數組
     */
    public static byte[] fileBuff(String filePath) throws IOException {
        File file = new File(filePath);
        long fileSize = file.length();
        if (fileSize > Integer.MAX_VALUE) {
            //System.out.println("file too big...");
            return null;
        }
        FileInputStream fi = new FileInputStream(file);
        byte[] file_buff = new byte[(int) fileSize];
        int offset = 0;
        int numRead = 0;
        while (offset < file_buff.length && (numRead = fi.read(file_buff, offset, file_buff.length - offset)) >= 0) {
            offset += numRead;
        }
        if (offset != file_buff.length) {
            throw new IOException("Could not completely read file " + file.getName());
        }
        fi.close();
        return file_buff;
    }


    public static void mergePdfFiles(String[] files, String savepath) {
        Document document = null;
        try {
            document = new Document(); 
            PdfCopy copy = new PdfCopy(document, new FileOutputStream(savepath));
            document.open();
            for (int i = 0; i < files.length; i++) {
                PdfReader reader = null;
                try {
                    reader = new PdfReader(files[i]);
                    int n = reader.getNumberOfPages();
                    for (int j = 1; j <= n; j++) {
                        document.newPage();
                        PdfImportedPage page = copy.getImportedPage(reader, j);
                        copy.addPage(page);
                    }
                } finally {
                    if (reader != null) {
                        reader.close();
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            
            if (document != null) {
                document.close();
            }

        }
    }
    
    public static void main(String[] args) throws Exception {    	
    	//Map中Key对应PDF表单中的fieldNames，Value则是你想填充的值
/*    	
    	var_year
    	var_month
    	var_day
    	var_name
    	var_id
    	var_birthday_year
    	var_birthday_month
    	var_birthday_day   */ 	
        HashMap map = new HashMap<String, String>();
        map.put("var_year","110");
        map.put("var_month","10");
        map.put("var_day","12");
        map.put("var_name","李述華");
        map.put("var_id","A220434569");
        map.put("var_birthday_year","61");
        map.put("var_birthday_month","06");
        map.put("var_birthday_day","22");
        map.put("var_personal","On");
        
        String sourceFile = "c:\\test.pdf"; //原文件路徑
        String targetFile = "result.pdf"; 	//目的文件路徑
      //  PdfUtils.genPdf(map,sourceFile,targetFile);    	
        genPdf(map,sourceFile,targetFile);   
    	
    }    	


        
}



