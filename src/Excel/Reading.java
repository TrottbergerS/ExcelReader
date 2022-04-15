package Excel;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.swing.plaf.synth.SynthOptionPaneUI;

import org.apache.commons.collections4.bag.SynchronizedSortedBag;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;




public class Reading {

	private static String subStr = "daa";

	private static boolean containsStr;
	
	public static void main(String[] args)  throws IOException 
	{  
 
		File file = new File("howtodoinjava_demo.xlsx");   //creating a new file instance  
		FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file 
		
		//creating Workbook instance that refers to .xlsx file  
		XSSFWorkbook wb = new XSSFWorkbook(fis);   
		XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object                 
		int rows=sheet.getLastRowNum()+1;
		int cols=sheet.getRow(1).getLastCellNum();
	
		String[][] intoo = new String[rows][cols];
	
		for (int r =0;r<rows; r++) {
			XSSFRow row=sheet.getRow(r);
			for(int c=0;c<cols;c++) {
			
				XSSFCell cell=row.getCell(c);
			
				intoo[r][c] = String.valueOf(cell);
				
			}
		}
		for (int a= 1;a< intoo.length;a++) {
			
			String[] parts = intoo[a][0].split("\\.|,");
			intoo[a][0] = parts[0];
			
			System.out.println("Nr. : " + intoo[a][0] + " Name: " + intoo[a][1] + " Link: " + intoo[a][2]);
		}	
		
		Scanner sc = new Scanner(System.in);
		
		int a1 = sc.nextInt();
		
		for (int a= 1;a< intoo.length;a++) {
			int nrtoint = Integer.parseInt(intoo[a][0]);
			// Abfrage ob Eintrag eine URL beinhaltet
			if (a1 == nrtoint ) {
				
				
				if (Pathtesting(intoo[a][2]) == true) {
					String explorerpfad = "explorer.exe " + intoo[a][2];
					System.out.println(explorerpfad);
					Runtime.getRuntime().exec(explorerpfad);
				}
				else {
					
					//System.out.println(intoo[a][b]);
					String celli = intoo[a][2];
					List<String> extractedUrls = extractUrls(celli);
					for (String url : extractedUrls){
					    //Wenn abfrage richtig dann wird in Browser geöffnet
					    ladeINet(url);
						}
				}
			}
		
		}	
	}

 // Abfrage ob URl
	public static List<String> extractUrls(String text)
	{
	    List<String> containedUrls = new ArrayList<String>();
	    String urlRegex = "((https?|ftp|gopher|telnet|file):((//)|(\\\\))+[\\w\\d:#@%/;$()~_?\\+-=\\\\\\.&]*)";
	    Pattern pattern = Pattern.compile(urlRegex, Pattern.CASE_INSENSITIVE);
	    Matcher urlMatcher = pattern.matcher(text);

	    while (urlMatcher.find())
	    {
	        containedUrls.add(text.substring(urlMatcher.start(0),
	                urlMatcher.end(0)));
	    }

	    return containedUrls;
	}
// Browser Öffnen
	public static void ladeINet(String seite) {
        try {
            Desktop.getDesktop().browse(new URI(seite));
        } catch (IOException e) {
            e.printStackTrace();
        } catch (URISyntaxException e) {
            e.printStackTrace();
        }
    } 
// Erkennung für Pfad
	public static boolean Pathtesting(String stringtest) {
		boolean back = false;
		String testing = "\\\\";
		

	if (stringtest.contains(testing)) {
		back = true;
	}
	    
		return back;
		
		
	}
}


