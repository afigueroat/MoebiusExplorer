package com.moebius.utilities;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;


public class ReadExcel {

	public static enum Option {ERROR, STATUS, VALUE, CLOSED, ORDER, CHECK, CAUSA};

	public static String FILE = "FILE";

	public static ArrayList<String> provideValue(int column) throws BiffException, IOException{
		Workbook workbook = Workbook.getWorkbook(new File(FILE));
		Sheet sheet = workbook.getSheet(0); 
		ArrayList<String> a = new ArrayList<String>();
		for (int i = 1; i < sheet.getRows() ; i++) {
			a.add(sheet.getCell(column,i).getContents());
		}
		return a;
	}
	public static String provideValue(int column, int row) throws BiffException, IOException{
		Workbook workbook = Workbook.getWorkbook(new File(FILE));
		Sheet sheet = workbook.getSheet(0); 
		return sheet.getCell(column,row).getContents();
	}
	public static void write(Option optionLabel, int fila, String value) throws BiffException, IOException, RowsExceededException, WriteException{
		Option option = optionLabel;
		Workbook target_workbook = Workbook.getWorkbook(new File(FILE));
		WritableWorkbook workbook = Workbook.createWorkbook(new File(FILE), target_workbook);
		target_workbook.close();
		WritableSheet sheet = workbook.getSheet(0);
		switch (option) {
		case ERROR : {
			sheet.addCell(new jxl.write.Label(13,fila,"ERROR"));
			break;
		}
		case CLOSED: {
			sheet.addCell(new jxl.write.Label(13,fila,"CLOSED"));
			break;
		}
		case STATUS: {
			sheet.addCell(new jxl.write.Label(1,fila,value));
			break;
		}
		case ORDER: {
			sheet.addCell(new jxl.write.Label(11,fila,value));
			break;
		}
		case CHECK: {
			sheet.addCell(new jxl.write.Label(12,fila,"OK"));
			break;
		}
		case CAUSA: {
			sheet.addCell(new jxl.write.Label(10, fila, value));
			break;
		}
		}
		workbook.write();
		workbook.close();
	}
}
