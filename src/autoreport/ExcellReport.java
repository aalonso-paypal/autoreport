

package autoreport;
import static autoreport.Autoreport.PorcentagemErro;
import static autoreport.Autoreport.Services;
import static autoreport.Autoreport.TopupServices;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import org.apache.pdfbox.cos.COSDocument;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.util.PDFTextStripper;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
//APACHE POI IMPORT
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
 
public class ExcellReport {
 
	// Extract text from PDF Document
	public static String pdftoText(String fileName) {
		PDFParser parser;
		String parsedText = null;;
		PDFTextStripper pdfStripper = null;
		PDDocument pdDoc = null;
		COSDocument cosDoc = null;
		File file = new File(fileName);
		if (!file.isFile()) {
			System.err.println("File " + fileName + " does not exist.");
			return null;
		}
		try {
			parser = new PDFParser(new FileInputStream(file));
		} catch (IOException e) {
			System.err.println("Unable to open PDF Parser. " + e.getMessage());
			return null;
		}
		try {
			parser.parse();
			cosDoc = parser.getDocument();
			pdfStripper = new PDFTextStripper();
			pdDoc = new PDDocument(cosDoc);
			pdfStripper.setStartPage(1);
			pdfStripper.setEndPage(5);
			parsedText = pdfStripper.getText(pdDoc);
		} catch (Exception e) {
			System.err.println("An exception occured in parsing the PDF Document."+e.getMessage());
		} finally {
			try {
				if (cosDoc != null)
					cosDoc.close();
				if (pdDoc != null)
					pdDoc.close();
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		return parsedText;
	}
        
        public static void FacebookPPWeb(String pathFile) {
            
            String response = pdftoText("c:/PAYPAL/PROJETOS/VIVO/Monthly Reports/"+pathFile);
            
            int limitator = response.indexOf("All rights reserved. For more information please visit www.gi-de.com", response.indexOf("Transaction Type Transaction Information Count"));
            System.out.println(response.subSequence(response.indexOf("Transaction Type Transaction Information Count"), limitator-2));
            
        }
        
        public static void TopupServices(String pathFile) {
            
            String response = pdftoText("c:/PAYPAL/PROJETOS/VIVO/Monthly Reports/"+pathFile);
            
            if(response.contains("No data found")){
                System.out.println("Nenhum erro reportado");
            }else{             
                response = pdftoText("c:/PAYPAL/PROJETOS/VIVO/Monthly Reports/"+pathFile);
            
                int limitator = response.indexOf("All rights reserved. For more information please visit www.gi-de.com", response.indexOf("Transaction Type Transaction Information Count"));
                System.out.println(response.subSequence(response.indexOf("Transaction Type Transaction Information Count"), limitator-2));

            }
        }
        
        public static float PorcentagemErro(String pathFile) {
            
            String response = pdftoText(pathFile);
            float retorno;
            float retornofinal;
            
            if(response.contains("No data found")){
                retornofinal=0.00f;
            }else{             
                response = pdftoText(pathFile);
                
                 String tempo = response.subSequence(response.indexOf("Success"),response.length()).toString();
                
                 String temp = tempo.subSequence(7,tempo.indexOf("Description")).toString();
                 
                 retorno = Float.parseFloat(temp.subSequence(temp.indexOf("Success")+9, temp.length()-3).toString());
                 retorno = (100.00f - retorno);
                 retornofinal = retorno;
                //int limitator = response.indexOf("All rights reserved. For more information please visit www.gi-de.com", response.indexOf("Transaction Type Transaction Information Count"));
                //System.out.println(response.subSequence(response.indexOf("Transaction Type Transaction Information Count"), limitator-2));
            }
            
            return retornofinal;
        }
        
        public static String Services(String pathFile) {
            String concat="";
            String response = pdftoText(pathFile);
            
            if(response.contains("No data found")){
                System.out.println("Nenhum erro reportado");
            }else{             
                response = pdftoText(pathFile);
                
                
                
                
                int limitator = response.indexOf("All rights reserved. For more information please visit www.gi-de.com", response.indexOf("Transaction Type Transaction Information Count"));
                response = response.subSequence(response.indexOf("Transaction Type Transaction Information Count"), limitator-2).toString();
                
                
                //<editor-fold defaultstate="collapsed" desc="BLOCO DE TRATAMENTO E LIMPEZA">
                
                response = response.replaceAll("\\s{2,}", " ");
                
                String[] temp = response.split("\\s", 2);
                
                response = temp[0]+"_"+temp[1];
                
                temp = response.split("\\s", 3);
                
                response = temp[0]+" "+temp[1]+"_"+temp[2];
                
                response = response.replace("c e", "ce");
                
                response = response.replace("v a", "va");
                
                response = response.replace("D e", "De");
                
                String[] broken = response.split("\\s");
                
                //</editor-fold>
                
                /*
                for (String broken1 : broken) {
                    System.out.println(broken1);
                }*/
                
                
                for(int i = 0; i < broken.length; i++){
                    if(i % 3 == 0 && i != 0){
                        concat = concat +"&"+broken[i];
                    }else{
                        if(i == 0){
                            concat = broken[i];
                        }else{
                            concat = concat +" "+ broken[i];
                        }
                    }
                }
                
                //System.out.println(concat);
                
            }
            if(concat.equals("")){
            concat+= " ";
            }
            return concat;
            //int limitator = response.indexOf("All rights reserved. For more information please visit www.gi-de.com", response.indexOf("Transaction Type Transaction Information Count"));
            //System.out.println(response.subSequence(response.indexOf("Transaction Type Transaction Information Count"), limitator-2));
        }
        
        public static void SheetGenerator(String Data,String sheetName,HSSFSheet sheet,HSSFWorkbook wb){
            
            String[] blocks = Data.split("%");
            
            //String[] rows = Data.split("&");
            
            //String[] hCell = rows[0].split("\\s");
            
            /*
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue(hCell[0]);
            header.createCell(1).setCellValue(hCell[1]);
            header.createCell(2).setCellValue(hCell[2]);
            
            String[] dCell = rows[i].split("\\s");
            Row dataRow = sheet.createRow(i);
            dataRow.createCell(0).setCellValue(dCell[0]);
            dataRow.createCell(1).setCellValue(dCell[1]);
            try{
                dataRow.createCell(2).setCellValue(dCell[2]);
                }catch(Exception e){
            
                }
            */
            int linhas=0;
            
            for(int i = 0; i < blocks.length; i++){
                
                String[] rows = blocks[i].split("&");
                
                String[] temp = rows[0].split("#");
                
                String[] reportTitle = temp[0].split("-");
                
                rows[0] = temp[1];
                
                Font f = wb.createFont();
                    

                    //set font 1 to 12 point type
                    f.setFontHeightInPoints((short) 12);
                    //make it blue
                    f.setColor(HSSFColor.WHITE.index);
                    // make it bold
                    //arial is the default font
                    f.setBoldweight(Font.BOLDWEIGHT_BOLD);
                    
                    CellStyle cs = wb.createCellStyle();
                    cs.setFont(f);
                    cs.setFillBackgroundColor(HSSFColor.BLUE.index);
                    cs.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                    
                    CellStyle cs2 = wb.createCellStyle();
                    cs2.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00%"));
                    cs2.setFont(f);
                    cs2.setFillBackgroundColor(HSSFColor.GREY_25_PERCENT.index);
                    cs2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                
                Row dataRow = sheet.createRow(linhas);
                Cell c = dataRow.createCell(0);
                c.setCellValue(reportTitle[0]);
                c.setCellStyle(cs);
                c = dataRow.createCell(1);
                c.setCellValue("Porcentagem de erro:");
                c.setCellStyle(cs);
                c = dataRow.createCell(2);
                c.setCellValue(reportTitle[1]);
                c.setCellStyle(cs2);
                linhas++;
                
                if(rows.length ==1){
                    dataRow = sheet.createRow(linhas);
                    dataRow.createCell(0).setCellValue("Sem erros relatados.");
                    linhas++;
                    
                }
            
                for (int j = 0; j < rows.length; j++ ){
                
                    String[] dCell = rows[j].split("\\s");
                    
                    dataRow = sheet.createRow(linhas);
                    try{
                    
                    
                    dataRow.createCell(0).setCellValue(dCell[0]);
                    dataRow.createCell(1).setCellValue(dCell[1]);
                    dataRow.createCell(2).setCellValue(dCell[2]);
                        }catch(Exception e){
                            dataRow.createCell(2).setCellValue("");
                        }
                    linhas++;
                    
                }
                    dataRow = sheet.createRow(linhas);
                    dataRow.createCell(0).setCellValue("");
                    dataRow.createCell(1).setCellValue("");
                    dataRow.createCell(2).setCellValue("");
                    linhas++;
                
            }
        
        }
        
        
        //<editor-fold defaultstate="collapsed" desc="Relatório">
        /*
        public static void main(String args[]){
            
            //Abre WorkSheet
            String sheetName = "Dezembro-2013";
            String Data = "";
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("Errors Report VIVO "+sheetName);
            
            //Inicia Leitura de pdfs e cria padrão de dados
            float perct;
             //Activate PIN Service Report
            System.out.println("-----------------------------------------------");
            perct = PorcentagemErro("Dezembro - 2013/ActivatePin_Service_Status-_December_20140101_024949.pdf");
            System.out.println("Activate PIN Service Report: "+perct);
            System.out.println("-----------------------------------------------");
            Data = "Activate PIN Service-"+perct+"#" +Services("Dezembro - 2013/ActivatePin_Service_Status-_December_20140101_024949.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //AddCard Service Report
            System.out.println("-----------------------------------------------");
            perct=PorcentagemErro("Dezembro - 2013/AddCard_Service_Status-_December_20140101_025313.pdf");
            System.out.println("AddCard Service Report: "+perct);
            System.out.println("-----------------------------------------------");
            Data += "%AddCard Service-"+perct+"#" +Services("Dezembro - 2013/AddCard_Service_Status-_December_20140101_025313.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //Delink Service Report
            System.out.println("-----------------------------------------------");
            perct=PorcentagemErro("Dezembro - 2013/Delink_Service_Status-_December_20140101_025324.pdf");
            System.out.println("Delink Service Report: "+perct);
            System.out.println("-----------------------------------------------");
            Data += "%Delink Service-"+perct+"#"+Services("Dezembro - 2013/Delink_Service_Status-_December_20140101_025324.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //RequestPayment Service Report
            System.out.println("-----------------------------------------------");
            perct=PorcentagemErro("Dezembro - 2013/RequestPayment_Service_Status-_December_20140101_024451.pdf");
            System.out.println("RequestPayment Service Report: "+perct);
            System.out.println("-----------------------------------------------");
            Data += "%RequestPayment Service-"+perct+"#"+Services("Dezembro - 2013/RequestPayment_Service_Status-_December_20140101_024451.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //ResetPIN Service Report
            System.out.println("-----------------------------------------------");
            perct=PorcentagemErro("Dezembro - 2013/ResetPIN_Service_Status-_December_20140101_024336.pdf");
            System.out.println("ResetPIN Service Report: "+perct);
            System.out.println("-----------------------------------------------");
            Data += "%ResetPIN Service-"+perct+"#"+Services("Dezembro - 2013/ResetPIN_Service_Status-_December_20140101_024336.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //SendPayment Service Report
            System.out.println("-----------------------------------------------");
            perct=PorcentagemErro("Dezembro - 2013/SendPayment_Service_Status-_December_20140101_024701.pdf");
            System.out.println("SendPayment Service Report: "+perct);
            System.out.println("-----------------------------------------------");
            Data += "%SendPayment Service-"+perct+"#"+Services("Dezembro - 2013/SendPayment_Service_Status-_December_20140101_024701.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //SignUp Service Report
            System.out.println("-----------------------------------------------");
            perct=PorcentagemErro("Dezembro - 2013/SignUp_Service_Status-_December_20140101_024825.pdf");
            System.out.println("SignUp Service Report: "+perct);
            System.out.println("-----------------------------------------------");
            Data += "%SignUp Service-"+perct+"#"+Services("Dezembro - 2013/SignUp_Service_Status-_December_20140101_024825.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //Facebook PP Topup Report
            System.out.println("-----------------------------------------------");
            perct=PorcentagemErro("Dezembro - 2013/FacebookPP_Web_Top-up_-_December_20140101_024621.pdf");
            System.out.println("Facebook PP Topup Report: "+perct);
            System.out.println("-----------------------------------------------");
            Data += "%Facebook PP Topup-"+perct+"#" +Services("Dezembro - 2013/FacebookPP_Web_Top-up_-_December_20140101_024621.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //Facebook VIVO Topup Report
            System.out.println("-----------------------------------------------");
            perct=PorcentagemErro("Dezembro - 2013/FacebookVivo_Web_Top-up_-_December_20140101_024535.pdf");
            System.out.println("Facebook VIVO Topup Report: "+perct);
            System.out.println("-----------------------------------------------");
            Data += "%Facebook VIVO Topup-"+perct+"#" +Services("Dezembro - 2013/FacebookVivo_Web_Top-up_-_December_20140101_024535.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //USSD Mobile Topup Report
            System.out.println("-----------------------------------------------");
            perct=PorcentagemErro("Dezembro - 2013/USSD_Mobile_Top-up_-_December_20140101_025554.pdf");
            System.out.println("USSD Mobile Topup Report: "+perct);
            System.out.println("-----------------------------------------------");
            Data += "%USSD Mobile Topup-"+perct+"#" +Services("Dezembro - 2013/USSD_Mobile_Top-up_-_December_20140101_025554.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //VivoWeb Web Topup Report
            System.out.println("-----------------------------------------------");
            perct=PorcentagemErro("Dezembro - 2013/VivoWeb_Web_Top-up_-_December_20140101_025132.pdf");
            System.out.println("VivoWeb Web Topup Report: "+perct);
            System.out.println("-----------------------------------------------");
            Data += "%VivoWeb Web Topup-"+perct+"#" +Services("Dezembro - 2013/VivoWeb_Web_Top-up_-_December_20140101_025132.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //VivoWebEn Web Topup Report
            System.out.println("-----------------------------------------------");
            perct=PorcentagemErro("Dezembro - 2013/VivoWebEn_Web_Top-up_-_December_20140101_025214.pdf");
            System.out.println("VivoWebEn Web Topup Report: "+perct);
            System.out.println("-----------------------------------------------");
            Data += "%VivoWebEn Web Topup-"+perct+"#" +Services("Dezembro - 2013/VivoWebEn_Web_Top-up_-_December_20140101_025214.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //VivoWebEs Web Topup Report
            System.out.println("-----------------------------------------------");
            perct=PorcentagemErro("Dezembro - 2013/VivoWebEs_Web_Top-up_-_December_20140101_024700.pdf");
            System.out.println("VivoWebEs Web Topup Report: "+perct);
            System.out.println("-----------------------------------------------");
            Data += "%VivoWebEs Web Topup-"+perct+"#" +Services("Dezembro - 2013/VivoWebEs_Web_Top-up_-_December_20140101_024700.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //VivoWebPtIntl Web Topup Report
            System.out.println("-----------------------------------------------");
            perct=PorcentagemErro("Dezembro - 2013/VivoWebPtIntl_Web_Top-up_-_December_20140101_025029.pdf");
            System.out.println("VivoWebPtIntl Web Topup Report: "+perct);
            System.out.println("-----------------------------------------------");
            Data += "%VivoWebPtIntl Web Topup-"+perct+"#" +Services("Dezembro - 2013/VivoWebPtIntl_Web_Top-up_-_December_20140101_025029.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
           
           
            System.out.println(Data);
            
            SheetGenerator(Data,sheetName,sheet,workbook);
            
            try {
                FileOutputStream out = new FileOutputStream(new File("C:\\PAYPAL\\PROJETOS\\VIVO\\Monthly Reports\\Consolidado\\Error Report Consolidado "+sheetName+".xls"));
                workbook.write(out);
                out.close();
                System.out.println("Excel written successfully..");

            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
                        
        }*/
 
}