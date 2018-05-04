package Operandsclass;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
public class ExcelReader {

static String operand;
static int valA;
static int valB;
static int Opvalcell ;
		// TODO Auto-generated method stub
		public static XSSFWorkbook wb;
		public static XSSFSheet wbsheet;
		public static XSSFRow row;
		public static XSSFCell cell;

		public static FileInputStream fis;
		public static FileOutputStream fileout;
		
		public static OutputStream fos;
		
		public static void main(String[] args) throws IOException {
			// TODO Auto-generated method stub
	              
			fis =new FileInputStream("C:\\Selenium\\udemy\\OperandCal.xlsx");

		      
		operand=getspecificCelldata(operand);
		System.out.println();
		
		fis.close();
		
}
	
		//function to call specific operand 
		 public static String getspecificCelldata(String operand) throws IOException
			{
				
				wb=new XSSFWorkbook(fis);
				
				 wbsheet=wb.getSheet("Operands");
				 int row=wbsheet.getLastRowNum();
				 System.out.println("Rows " +row);
				
			int col=wbsheet.getRow(0).getLastCellNum();	
			
			System.out.println("Column " + col);
			
			for(int i=1;i<=row;i++)
			{
				XSSFCell opcell=wbsheet.getRow(i).getCell(0);
				XSSFCell Acell=wbsheet.getRow(i).getCell(1);//Cells under ColA
				XSSFCell Acel2=wbsheet.getRow(i).getCell(2);//Cells under ColB
				
				 operand =opcell.getStringCellValue();
				
				 valA=(int) Acell.getNumericCellValue();
				 valB=(int) Acel2.getNumericCellValue();
				 
				
				 Row rowcal = wbsheet.getRow(i);
				 Cell cell = rowcal.getCell(3);
				 if (cell == null)
				     cell = rowcal.createCell(3);
				 cell.setCellType(Cell.CELL_TYPE_NUMERIC);

				 
				
				System.out.println("Operand " + operand);
				if(operand.equalsIgnoreCase("+"))
				{
					System.out.println("Call plus operand");
					Add add = new Add();
				
					 int c = add.calculate(operand,valA,valB);
					System.out.println("Final value : " +c);
				
				cell.setCellValue(c);
				 int valC =(int) cell.getNumericCellValue();

				System.out.println("Value of cell after calcultion : " + " of .."+ operand + "..  " +valC);
			fileout = new FileOutputStream("C:\\Selenium\\udemy\\Cal.xlsx");
		            wb.write(fileout);
		           
		            fis.close();

				}
				
				
				
				else if(operand.equalsIgnoreCase("-"))
				{
					System.out.println("Call minus operand");
					
					Minus min = new Minus();
					
					 int c = min.calculate(operand,valA,valB);
					System.out.println("Final value : " +c);
					 cell.setCellValue(c);
				
					 int valC =(int) cell.getNumericCellValue();

						System.out.println("Value of cell after calcultion : " + " of .."+ operand + "..  " +valC);
						fileout = new FileOutputStream("C:\\Selenium\\udemy\\Cal.xlsx");
					            wb.write(fileout);
					            fis.close();
				}
				
				else if(operand.equalsIgnoreCase("*"))
				{
					System.out.println("Call multiplication operand");
					
					
					Multiplication mul = new Multiplication();
					
					 int c = mul.calculate(operand,valA,valB);
					System.out.println("Final value : " +c);
					cell.setCellValue(c);
					
					 int valC =(int) cell.getNumericCellValue();

						System.out.println("Value of cell after calcultion : " + " of .."+ operand + "..  " +valC);
						fileout = new FileOutputStream("C:\\Selenium\\udemy\\Cal.xlsx");
					            wb.write(fileout);
					            fis.close();
                            }
				
				else if(operand.equalsIgnoreCase("/"))
				{
					System.out.println("Call division operand");
					
					Division div = new Division();
					
					 int c = div.calculate(operand,valA,valB);
					System.out.println("Final value : " +c);
					cell.setCellValue(c);
					
					 int valC =(int) cell.getNumericCellValue();

						System.out.println("Value of cell after calcultion : " + " of .."+ operand + "..  " +valC);
						fileout = new FileOutputStream("C:\\Selenium\\udemy\\Cal.xlsx");
					            wb.write(fileout);
					            fis.close();
				}
				
				else if(operand.equalsIgnoreCase("%"))
				{
					System.out.println("Call  percentage operand");

					
					Percentage per = new Percentage();
					
					 int c = per.calculate(operand,valA,valB);
					System.out.println("Final value : " +c);
					cell.setCellValue(c);
					
					 int valC =(int) cell.getNumericCellValue();

						System.out.println("Value of cell after calcultion : " + " of .."+ operand + "..  " +valC);
						fileout = new FileOutputStream("C:\\Selenium\\udemy\\Cal.xlsx");
					            wb.write(fileout);
					            fis.close();
				}
				
				else {
					try {
					
						System.out.println("Not a valid operand ");
					}
					catch(Exception e)
					{
						e.printStackTrace();
					}
				}
				
				
			}
			return operand;
			
			
			}
		
		
	
	}


//