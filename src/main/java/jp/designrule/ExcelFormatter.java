package jp.designrule;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.FileOutputStream;
import java.io.FileInputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFormatter {

  public static void main(String... args) throws IOException, InvalidFormatException {
    File file = new File(args[0]);
    Workbook workbook = null;
    try(InputStream is = new FileInputStream(file)) {
        workbook = WorkbookFactory.create(is);
      }catch(Exception e) {
      e.printStackTrace();
      System.exit(-1);
    }

    Sheet sheet = workbook.getSheet("Sheet1");
    
    //B31の参照先を書き換え
    replaceFormula(sheet, "B31", "'01_0表紙'!t28");

    //上書き保存
    try (FileOutputStream fos = new FileOutputStream(file)) {
        workbook.write(fos);
      }catch(Exception e) {
      e.printStackTrace();
      System.exit(-1);      
    }
  }

  public static void replaceFormula(Sheet sheet, String cellName, String replacement) {
    CellReference reference = new CellReference(cellName); // A1形式
    Row row = sheet.getRow(reference.getRow());
    Cell cell = row.getCell(reference.getCol());
    cell.setCellFormula(replacement);
  }
  
}
