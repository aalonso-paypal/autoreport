

package autoreport;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.pdfbox.cos.COSDocument;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.util.PDFTextStripper;
 
public class Autoreport {
 
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
        
        public static String PorcentagemErro(String pathFile) {
            
            String response = pdftoText("c:/PAYPAL/PROJETOS/VIVO/Monthly Reports/"+pathFile);
            float retorno;
            String retornofinal="";
            
            if(response.contains("No data found")){
                retornofinal="0.00% de erro";
            }else{             
                response = pdftoText("c:/PAYPAL/PROJETOS/VIVO/Monthly Reports/"+pathFile);
                
                 String tempo = response.subSequence(response.indexOf("Success"),response.length()).toString();
                
                 String temp = tempo.subSequence(7,tempo.indexOf("Description")).toString();
                 
                 retorno = Float.parseFloat(temp.subSequence(temp.indexOf("Success")+9, temp.length()-3).toString());
                 retorno = (100.00f - retorno);
                 retornofinal = retorno+"% de erro";
                //int limitator = response.indexOf("All rights reserved. For more information please visit www.gi-de.com", response.indexOf("Transaction Type Transaction Information Count"));
                //System.out.println(response.subSequence(response.indexOf("Transaction Type Transaction Information Count"), limitator-2));
            }
            
            return retornofinal;
        }
        
        public static void Services(String pathFile) {
            
            String response = pdftoText("c:/PAYPAL/PROJETOS/VIVO/Monthly Reports/"+pathFile);
            
            if(response.contains("No data found")){
                System.out.println("Nenhum erro reportado");
            }else{             
                response = pdftoText("c:/PAYPAL/PROJETOS/VIVO/Monthly Reports/"+pathFile);
            
                int limitator = response.indexOf("All rights reserved. For more information please visit www.gi-de.com", response.indexOf("Transaction Type Transaction Information Count"));
                System.out.println(response.subSequence(response.indexOf("Transaction Type Transaction Information Count"), limitator-2));

            }
            
            //int limitator = response.indexOf("All rights reserved. For more information please visit www.gi-de.com", response.indexOf("Transaction Type Transaction Information Count"));
            //System.out.println(response.subSequence(response.indexOf("Transaction Type Transaction Information Count"), limitator-2));
        }
        
        /*
        //<editor-fold defaultstate="collapsed" desc="Relatório">
        public static void main(String args[]){
            //Inicio de bloco de Serviços
            System.out.println("_______________________________________________");
            System.out.println("                  SERVICES");
            System.out.println("_______________________________________________");
            System.out.println("\n");
            
            //Activate PIN Service Report
            System.out.println("-----------------------------------------------");
            System.out.println("Activate PIN Service Report: "+PorcentagemErro("Dezembro - 2013/ActivatePin_Service_Status-_December_20140101_024949.pdf"));
            System.out.println("-----------------------------------------------");
            Services("Dezembro - 2013/ActivatePin_Service_Status-_December_20140101_024949.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //AddCard Service Report
            System.out.println("-----------------------------------------------");
            System.out.println("AddCard Service Report: "+PorcentagemErro("Dezembro - 2013/AddCard_Service_Status-_December_20140101_025313.pdf"));
            System.out.println("-----------------------------------------------");
            Services("Dezembro - 2013/AddCard_Service_Status-_December_20140101_025313.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //Delink Service Report
            System.out.println("-----------------------------------------------");
            System.out.println("Delink Service Report: "+PorcentagemErro("Dezembro - 2013/Delink_Service_Status-_December_20140101_025324.pdf"));
            System.out.println("-----------------------------------------------");
            Services("Dezembro - 2013/Delink_Service_Status-_December_20140101_025324.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //RequestPayment Service Report
            System.out.println("-----------------------------------------------");
            System.out.println("RequestPayment Service Report: "+PorcentagemErro("Dezembro - 2013/RequestPayment_Service_Status-_December_20140101_024451.pdf"));
            System.out.println("-----------------------------------------------");
            Services("Dezembro - 2013/RequestPayment_Service_Status-_December_20140101_024451.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //ResetPIN Service Report
            System.out.println("-----------------------------------------------");
            System.out.println("ResetPIN Service Report: "+PorcentagemErro("Dezembro - 2013/ResetPIN_Service_Status-_December_20140101_024336.pdf"));
            System.out.println("-----------------------------------------------");
            Services("Dezembro - 2013/ResetPIN_Service_Status-_December_20140101_024336.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //SendPayment Service Report
            System.out.println("-----------------------------------------------");
            System.out.println("SendPayment Service Report: "+PorcentagemErro("Dezembro - 2013/SendPayment_Service_Status-_December_20140101_024701.pdf"));
            System.out.println("-----------------------------------------------");
            Services("Dezembro - 2013/SendPayment_Service_Status-_December_20140101_024701.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //SignUp Service Report
            System.out.println("-----------------------------------------------");
            System.out.println("SignUp Service Report: "+PorcentagemErro("Dezembro - 2013/SignUp_Service_Status-_December_20140101_024825.pdf"));
            System.out.println("-----------------------------------------------");
            Services("Dezembro - 2013/SignUp_Service_Status-_December_20140101_024825.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //Inicio de bloco de Serviços
            System.out.println("_______________________________________________");
            System.out.println("                  TOPUP");
            System.out.println("_______________________________________________");
            System.out.println("\n");
            
            //Facebook PP Topup Report
            System.out.println("-----------------------------------------------");
            System.out.println("Facebook PP Topup Report: "+PorcentagemErro("Dezembro - 2013/FacebookPP_Web_Top-up_-_December_20140101_024621.pdf"));
            System.out.println("-----------------------------------------------");
            TopupServices("Dezembro - 2013/FacebookPP_Web_Top-up_-_December_20140101_024621.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //Facebook VIVO Topup Report
            System.out.println("-----------------------------------------------");
            System.out.println("Facebook VIVO Topup Report: "+PorcentagemErro("Dezembro - 2013/FacebookVivo_Web_Top-up_-_December_20140101_024535.pdf"));
            System.out.println("-----------------------------------------------");
            TopupServices("Dezembro - 2013/FacebookVivo_Web_Top-up_-_December_20140101_024535.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //USSD Mobile Topup Report
            System.out.println("-----------------------------------------------");
            System.out.println("USSD Mobile Topup Report: "+PorcentagemErro("Dezembro - 2013/USSD_Mobile_Top-up_-_December_20140101_025554.pdf"));
            System.out.println("-----------------------------------------------");
            TopupServices("Dezembro - 2013/USSD_Mobile_Top-up_-_December_20140101_025554.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //VivoWeb Web Topup Report
            System.out.println("-----------------------------------------------");
            System.out.println("VivoWeb Web Topup Report: "+PorcentagemErro("Dezembro - 2013/VivoWeb_Web_Top-up_-_December_20140101_025132.pdf"));
            System.out.println("-----------------------------------------------");
            TopupServices("Dezembro - 2013/VivoWeb_Web_Top-up_-_December_20140101_025132.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //VivoWebEn Web Topup Report
            System.out.println("-----------------------------------------------");
            System.out.println("VivoWebEn Web Topup Report: "+PorcentagemErro("Dezembro - 2013/VivoWebEn_Web_Top-up_-_December_20140101_025214.pdf"));
            System.out.println("-----------------------------------------------");
            TopupServices("Dezembro - 2013/VivoWebEn_Web_Top-up_-_December_20140101_025214.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //VivoWebEs Web Topup Report
            System.out.println("-----------------------------------------------");
            System.out.println("VivoWebEs Web Topup Report: "+PorcentagemErro("Dezembro - 2013/VivoWebEs_Web_Top-up_-_December_20140101_024700.pdf"));
            System.out.println("-----------------------------------------------");
            TopupServices("Dezembro - 2013/VivoWebEs_Web_Top-up_-_December_20140101_024700.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
            //VivoWebPtIntl Web Topup Report
            System.out.println("-----------------------------------------------");
            System.out.println("VivoWebPtIntl Web Topup Report: "+PorcentagemErro("Dezembro - 2013/VivoWebPtIntl_Web_Top-up_-_December_20140101_025029.pdf")+"");
            System.out.println("-----------------------------------------------");
            TopupServices("Dezembro - 2013/VivoWebPtIntl_Web_Top-up_-_December_20140101_025029.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
        }
//</editor-fold>
        */
        /*
        //<editor-fold defaultstate="collapsed" desc="Teste">
        public static void main(String args[]){
            
            //Facebook PP Topup Report
            System.out.println("-----------------------------------------------");
            System.out.println("Facebook PP Topup Report:");
            System.out.println("-----------------------------------------------");
            PorcentagemErro("Dezembro - 2013/FacebookPP_Web_Top-up_-_December_20140101_024621.pdf");
            System.out.println("_______________________________________________\n");
            System.out.println("\n");
            
        }
//</editor-fold>*/
 
}