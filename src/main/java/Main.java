//import Excel.ReadFromExcel;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

public class Main {
    public static void main(String[] args){
        try {
            FileInputStream inputStream = new FileInputStream("CS.xlsx");//создаём поток ввода из экселя
            //FileOutputStream writeIntoWord = new FileOutputStream("output.docx");//поток для вывода в word
            try {
                //открываем word
                //XWPFDocument wordDocument = new XWPFDocument();

                //открываем эксель
                // Get the workbook instance for XLS file
                XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
                // Get first sheet from the workbook
                XSSFSheet sheet = workbook.getSheetAt(0);
                // Get iterator to all the rows in current sheet
                Iterator<Row> rowIterator = sheet.iterator();

                while (rowIterator.hasNext()) {//пока есть следующая строка
                    Row row = rowIterator.next();
                    // Get iterator to all cells of current row
                    Iterator<Cell> cellIterator = row.cellIterator();

                    while (cellIterator.hasNext()) {//пока в строке есть следующая ячейка
                        Cell cell = cellIterator.next();

                        // Change to getCellType() if using POI 4.x
                        //CellType cellType = cell.getCellTypeEnum();
                        CellType cellType = cell.getCellType();//узнаём формат ячейки
                        switch (cellType) {//в зависимости от формата ячейки выбираем, что печатать
                            case _NONE:
                                System.out.print("");
                                System.out.print("\t");
                                break;
                            case BOOLEAN:
                                System.out.print(cell.getBooleanCellValue());
                                System.out.print("\t");
                                break;
                            case BLANK:
                                System.out.print("");
                                System.out.print("\t");
                                break;
                            case FORMULA:
                                // Formula
                                System.out.print(cell.getCellFormula());
                                System.out.print("\t");

                                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                                // Print out value evaluated by formula
                                System.out.print(evaluator.evaluate(cell).getNumberValue());
                                break;
                            case NUMERIC:
                                System.out.print(cell.getNumericCellValue());
                                System.out.print("\t");
                                break;
                            case STRING:
                                System.out.print(cell.getStringCellValue());
                                System.out.print("\t");
                                break;
                            case ERROR:
                                System.out.print("!");
                                System.out.print("\t");
                                break;
                        }
                        //wordDocument.write(writeIntoWord);
                        writeIntoWord.write(79);
                    }
                    System.out.println("");
                }
            } catch(IOException e){
                System.out.println("error");
            }

        }

        catch(FileNotFoundException e){
            System.out.println("file not found");
        }

    }
}
