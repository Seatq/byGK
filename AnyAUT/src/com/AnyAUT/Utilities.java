package com.AnyAUT;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.IndexedColors;


public class Utilities {
	static String filePath;

	//Purpose: This method identifies absolute path of the current workspace
	public static String fileAbsolutePath() {
		
		//fetch the absolute location of the current workspace
		filePath = new java.io.File("").getAbsolutePath();
		 
		//returns the absolute path 
		 return filePath+"/";
	}//end of fileAbsolutePath
	
	// Method to read XL
	// Purpose: This method reads an excel based on the file path and sheet name provided
	public static String[][] readXL(String fPath, String fSheet) throws Exception {
		// Inputs : XL Path and XL Sheet name
		// Output : 
		String[][] xData;  
		int xRows, xCols;
		DataFormatter dataFormatter = new DataFormatter();
		String cellValue;
		File myxl = new File(fPath);                                
		FileInputStream myStream = new FileInputStream(myxl);                                
		HSSFWorkbook myWB = new HSSFWorkbook(myStream);                                
		HSSFSheet mySheet = myWB.getSheet(fSheet);                                 
		xRows = mySheet.getLastRowNum()+1;                                
		xCols = mySheet.getRow(0).getLastCellNum();   
		xData = new String[xRows][xCols];   

		for (int i = 0; i < xRows; i++) {  
			if (mySheet.getRow(i) == null) continue; //Updated by Ragi - To avoid null pointer exception in empty rows
			HSSFRow row = mySheet.getRow(i);
			for (int j = 0; j < xCols; j++) { 
				cellValue = "-";
				cellValue = dataFormatter.formatCellValue(row.getCell(j));
				if (cellValue!=null) {
					xData[i][j] = cellValue; 
				}  
			}        
		}    
		mySheet = null;
		myxl = null; 
		myWB = null;
		return xData;
	}


	public static CellStyle headingcellformat(HSSFWorkbook fwb)
	{
		CellStyle style = fwb.createCellStyle();
		HSSFFont font =fwb.createFont();
		font.setBold(true);
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setFont(font);
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setBorderLeft(CellStyle.BORDER_THIN);
		return style;
	}
			
	public static CellStyle datacellformat(HSSFWorkbook fwb)
	{
		CellStyle style = fwb.createCellStyle();
		HSSFFont font =fwb.createFont();
		font.setFontName("Courier New");
		style.setFillForegroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setFont(font);
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setBorderLeft(CellStyle.BORDER_THIN);
		return style;
	}		

	public static CellStyle passedcellformat(HSSFWorkbook fwb)
	{
		CellStyle style = fwb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setBorderLeft(CellStyle.BORDER_THIN);
		return style;
	}	

	public static CellStyle failedcellformat(HSSFWorkbook fwb)
	{
		CellStyle style = fwb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.RED.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setBorderLeft(CellStyle.BORDER_THIN);
		return style;
	}		
		
	public static void writeXLSheets(String sPath, String iSheet,int sheetIndex, String[][] xData)
			throws Exception
	{
		HSSFWorkbook  wb,newWB;
		HSSFSheet osheet;
		File outFile = new File(sPath);
		if ((outFile.isFile() == true)&&(outFile.exists()==true))
		{
			final InputStream is = new FileInputStream(outFile);
			try {
					wb = new HSSFWorkbook(is);
					if((wb.getNumberOfSheets())== sheetIndex)
					wb.createSheet();
		
					if((wb.getSheetName(sheetIndex)).equalsIgnoreCase(iSheet))
					{
						wb.removeSheetAt(sheetIndex);
						osheet=wb.createSheet(iSheet);
						wb.setSheetOrder(iSheet, sheetIndex);
						CellStyle style=headingcellformat(wb);
						CellStyle style1=datacellformat(wb);
						CellStyle style2=passedcellformat(wb);
						CellStyle style3=failedcellformat(wb);
			
						int xR_TS = xData.length;
					       int xC_TS = xData[0].length;
					   	for (int myrow = 0; myrow < xR_TS; myrow++)
					   	{
						       HSSFRow row = osheet.createRow(myrow);
						       for (int mycol = 0; mycol < xC_TS; mycol++)
						       {
						       	HSSFCell cell = row.createCell(mycol);
						       	cell.setCellType(HSSFCell.CELL_TYPE_STRING);
						       	cell.setCellValue(xData[myrow][mycol]);
						       	if(myrow==0)
						       	{cell.setCellStyle(style);}
						       	else{cell.setCellStyle(style1);}
						     	////////////////////////////////////////////
						       if ((xData[myrow][mycol]).equals("Pass"))
						       	{cell.setCellStyle(style2);}
						       	else if (xData[myrow][mycol].equals("Fail"))
						       	{cell.setCellStyle(style3);}
						       /////////////////////////////////////////////// 
						       }
						       FileOutputStream fOut = new FileOutputStream(outFile);
						       wb.write(fOut);
						       fOut.flush();
						       fOut.close();
					   	}//outerfor 
			
					}//end if sheet exists
					else 
					{
						osheet=wb.createSheet(iSheet);
						wb.setSheetOrder(iSheet, sheetIndex);
						CellStyle style=headingcellformat(wb);
						CellStyle style1=datacellformat(wb);
						int xR_TS = xData.length;
					    int xC_TS = xData[0].length;
					   	for (int myrow = 0; myrow < xR_TS; myrow++)
					   	{
					   		HSSFRow row = osheet.createRow(myrow);
					   		for (int mycol = 0; mycol < xC_TS; mycol++)
					   		{
						       	HSSFCell cell = row.createCell(mycol);
						       	cell.setCellType(HSSFCell.CELL_TYPE_STRING);
						       	cell.setCellValue(xData[myrow][mycol]);
						       	if(myrow==0)
							       	cell.setCellStyle(style);
							     else
							       	cell.setCellStyle(style1);
					   		}
					   		if (wb.getSheetName(wb.getNumberOfSheets()-1).equalsIgnoreCase("Sheet1")) 
					   			wb.removeSheetAt(wb.getNumberOfSheets()-1);
							FileOutputStream fOut = new FileOutputStream(outFile);
						    wb.write(fOut);
						    fOut.flush();
						    fOut.close();
					   	} 
					}
				} 
			finally 
			{ 
				is.close();
			}
		}//end of first if file exists
		else if (outFile.isFile() == false)
		{
			newWB = new HSSFWorkbook();
			HSSFSheet newsheet = newWB.createSheet(iSheet);
			CellStyle style=headingcellformat(newWB);
			CellStyle style1=datacellformat(newWB);
			int xR_TS = xData.length;
			int xC_TS = xData[0].length;
		   	for (int myrow = 0; myrow < xR_TS; myrow++) 
		   	{
			       HSSFRow row = newsheet.createRow(myrow);
			       for (int mycol = 0; mycol < xC_TS; mycol++) 
			       {
				       	HSSFCell cell = row.createCell(mycol);
				       	cell.setCellType(HSSFCell.CELL_TYPE_STRING);
				       	cell.setCellValue(xData[myrow][mycol]);
				       	if(myrow == 0)
					       	cell.setCellStyle(style);
					       	else
					       	cell.setCellStyle(style1); 
			       }
			       if (newWB.getSheetName(newWB.getNumberOfSheets()-1).equalsIgnoreCase("Sheet1")) 
			    	   newWB.removeSheetAt(newWB.getNumberOfSheets()-1);
			       FileOutputStream fOut = new FileOutputStream(outFile);
			       newWB.write(fOut);
			       fOut.flush();
			       fOut.close();
		   	}
		}
		 

	}//end of writeXLSheets
	public static void mailRunResult(String xlPath) throws IOException {
		// authentication info
		final String username = "myaut.results@gmail.com";
		final String password = "Anyaut123";
		String fromEmail = "myaut.results@gmail.com";
		System.out.println("**EMAIL RETREIVED FROM RUNNER IS " + UICode.strEmail + "****");
		String toEmail = UICode.strEmail;

		Properties properties = new Properties();
		properties.put("mail.smtp.auth", "true");
		properties.put("mail.smtp.starttls.enable", "true");

		properties.put("mail.smtp.host", "smtp.gmail.com");
		properties.put("mail.smtp.port", "587");

		Session session = Session.getInstance(properties, new javax.mail.Authenticator() {
			protected PasswordAuthentication getPasswordAuthentication() {
				return new PasswordAuthentication(username, password);
			}
		});
		// Start our mail message
		MimeMessage msg = new MimeMessage(session);
		String vText = "Hello There!..This is the message to check email has been sent using Java!! ";
		try {
			msg.setFrom(new InternetAddress(fromEmail));
			msg.addRecipient(Message.RecipientType.TO, new InternetAddress(toEmail));
			msg.setSubject("Sending email from Eclipse");

			Multipart emailContent = new MimeMultipart();

			// Text body part
			MimeBodyPart textBodyPart = new MimeBodyPart();
			textBodyPart.setText(vText);

			// Attachment body part.
			MimeBodyPart excelAttachment = new MimeBodyPart();
			// excelAttachment.attachFile("C:\\Users\\T430S\\Documents\\SampleAttachment.xls");
			excelAttachment.attachFile(xlPath);

			// Attach body parts
			emailContent.addBodyPart(textBodyPart);
			emailContent.addBodyPart(excelAttachment);
			// Attach multipart to message
			msg.setContent(emailContent);
			Transport.send(msg);
			System.out.println("Sent message");
		} catch (MessagingException e) {
			e.printStackTrace();
		}

	}
}//end of class
