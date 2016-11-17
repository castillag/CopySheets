/*
 * Este Programa copia hojas de ficheros excel en otro fichero excel
 * Hay que seleccionar los ficheros donde estan las hojas que se quieren
 * copiar, y el fichero nuevo a donde compiarlas
 * Hay que indicar el numero de hoja que se quiere copiar
 */

package copysheets;
/**
* all credits go to
* http://www.coderanch.com/t/420958/open-source/Copying-sheet-excel-file-another
* in teh Forum: Other Open Source Projects
* and adapted by Guillermo Castilla
**/
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CopySheets {
    
/** 
 * @param newSheet the sheet to create from the copy. 
 * @param sheet the sheet to copy. 
 */  
public static void copySheets(Sheet newSheet, Sheet sheet){     
    copySheets(newSheet, sheet, true);     
}     

/** 
 * @param newSheet the sheet to create from the copy. 
 * @param sheet the sheet to copy. 
 * @param copyStyle true copy the style. 
 */  
public static void copySheets(Sheet newSheet, Sheet sheet, boolean copyStyle){     
    int maxColumnNum = 0;     
    Map<Integer, CellStyle> styleMap = (copyStyle) ? new HashMap<Integer, CellStyle>() : null;     
    for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {     
        Row srcRow = sheet.getRow(i);     
        Row destRow = newSheet.createRow(i);     
        if (srcRow != null) {     
            copyRow(sheet, newSheet, srcRow, destRow, styleMap);     
            if (srcRow.getLastCellNum() > maxColumnNum) {     
                maxColumnNum = srcRow.getLastCellNum();     
            }     
        }     
    }     
    for (int i = 0; i <= maxColumnNum; i++) {     
        newSheet.setColumnWidth(i, sheet.getColumnWidth(i));     
    }     
}     

/** 
 * @param srcSheet the sheet to copy. 
 * @param destSheet the sheet to create. 
 * @param srcRow the row to copy. 
 * @param destRow the row to create. 
 * @param styleMap - 
 */  
public static void copyRow(Sheet srcSheet, Sheet destSheet, Row srcRow, Row destRow, Map<Integer, CellStyle> styleMap) {     
    // manage a list of merged zone in order to not insert two times a merged zone  
  Set<CellRangeAddressWrapper> mergedRegions = new TreeSet<CellRangeAddressWrapper>();     
    destRow.setHeight(srcRow.getHeight());     
    // reckoning delta rows  
    int deltaRows = destRow.getRowNum()-srcRow.getRowNum();  
    // pour chaque row  
    for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {     
        Cell oldCell = srcRow.getCell(j);   // ancienne cell  
        Cell newCell = destRow.getCell(j);  // new cell   
        if (oldCell != null) {     
            if (newCell == null) {     
                newCell = destRow.createCell(j);     
            }     
            // copy chaque cell  
            copyCell(oldCell, newCell, styleMap);     
            // copy les informations de fusion entre les cellules  
            //System.out.println("row num: " + srcRow.getRowNum() + " , col: " + (short)oldCell.getColumnIndex());  
            CellRangeAddress mergedRegion = getMergedRegion(srcSheet, srcRow.getRowNum(), (short)oldCell.getColumnIndex());     

            if (mergedRegion != null) {   
              //System.out.println("Selected merged region: " + mergedRegion.toString());  
              CellRangeAddress newMergedRegion = new CellRangeAddress(mergedRegion.getFirstRow()+deltaRows, mergedRegion.getLastRow()+deltaRows, mergedRegion.getFirstColumn(),  mergedRegion.getLastColumn());  
                //System.out.println("New merged region: " + newMergedRegion.toString());  
                CellRangeAddressWrapper wrapper = new CellRangeAddressWrapper(newMergedRegion);  
                if (isNewMergedRegion(wrapper, mergedRegions)) {  
                    mergedRegions.add(wrapper);  
                    destSheet.addMergedRegion(wrapper.range);     
                }     
            }     
        }     
    }                
}    

/** 
 * @param oldCell 
 * @param newCell 
 * @param styleMap 
 */  
public static void copyCell(Cell oldCell, Cell newCell, Map<Integer, CellStyle> styleMap) {     
    if(styleMap != null) {     
        if(oldCell.getSheet().getWorkbook() == newCell.getSheet().getWorkbook()){     
            newCell.setCellStyle(oldCell.getCellStyle());     
        } else{     
            int stHashCode = oldCell.getCellStyle().hashCode();     
            CellStyle newCellStyle = styleMap.get(stHashCode);     
            if(newCellStyle == null){     
                newCellStyle = newCell.getSheet().getWorkbook().createCellStyle();     
                newCellStyle.cloneStyleFrom(oldCell.getCellStyle());     
                styleMap.put(stHashCode, newCellStyle);     
            }     
            newCell.setCellStyle(newCellStyle);     
        }     
    }     
    switch(oldCell.getCellType()) {     
        case Cell.CELL_TYPE_STRING:     
            newCell.setCellValue(oldCell.getStringCellValue());     
            break;     
      case Cell.CELL_TYPE_NUMERIC:     
            newCell.setCellValue(oldCell.getNumericCellValue());     
            break;     
        case Cell.CELL_TYPE_BLANK:     
            newCell.setCellType(XSSFCell.CELL_TYPE_BLANK);     
            break;     
        case Cell.CELL_TYPE_BOOLEAN:     
            newCell.setCellValue(oldCell.getBooleanCellValue());     
            break;     
        case Cell.CELL_TYPE_ERROR:     
            newCell.setCellErrorValue(oldCell.getErrorCellValue());     
            break;     
        case Cell.CELL_TYPE_FORMULA:     
            newCell.setCellFormula(oldCell.getCellFormula());     
            break;     
        default:     
            break;     
    }     

}     

/** 
 * Récupère les informations de fusion des cellules dans la sheet source pour les appliquer 
 * à la sheet destination... 
 * Récupère toutes les zones merged dans la sheet source et regarde pour chacune d'elle si 
 * elle se trouve dans la current row que nous traitons. 
 * Si oui, retourne l'objet CellRangeAddress. 
 *  
 * @param sheet the sheet containing the data. 
 * @param rowNum the num of the row to copy. 
 * @param cellNum the num of the cell to copy. 
 * @return the CellRangeAddress created. 
 */  
public static CellRangeAddress getMergedRegion(Sheet sheet, int rowNum, short cellNum) {     
    for (int i = 0; i < sheet.getNumMergedRegions(); i++) {   
        CellRangeAddress merged = sheet.getMergedRegion(i);     
        if (merged.isInRange(rowNum, cellNum)) {     
            return merged;     
        }     
    }     
    return null;     
}     

/** 
 * Check that the merged region has been created in the destination sheet. 
 * @param newMergedRegion the merged region to copy or not in the destination sheet. 
 * @param mergedRegions the list containing all the merged region. 
 * @return true if the merged region is already in the list or not. 
 */  
private static boolean isNewMergedRegion(CellRangeAddressWrapper newMergedRegion, Set<CellRangeAddressWrapper> mergedRegions) {  
  return !mergedRegions.contains(newMergedRegion);     
}     


  

//------------
/**
 * 
 * @param files ArrayList con los Ficheros de los que hay que copiar la hoja 5
 * @param f Fichero en el que hay que copiarlos
 * @throws FileNotFoundException
 * @throws IOException 
 */
public static void inicio(File[] files, String fname, int nhoja ) throws FileNotFoundException, IOException  {
        // TODO code application logic here
  //  try{
        Workbook book = null;
        //------------------Datos de Entrada a modificar -------------------------
        int numero_de_ficheros = files.length;
        int numero_hoja = nhoja-1;
        FileInputStream [] hoja = new FileInputStream[numero_de_ficheros] ;        
        File f = new File(fname);
        if (f.exists()) {
                if (fname.endsWith("xlsx")) {
                    book = new XSSFWorkbook(new FileInputStream(fname));
               //      book = WorkbookFactory.create(new FileInputStream(fname));
                } else {
                    try {
                        throw new Exception("Debe ser un fichero Excel \"xlsx\"");
                    } catch (Exception ex) {
                        Logger.getLogger(CopySheets.class.getName()).log(Level.SEVERE, null, ex);
                    }
                }
            } else {
                f.createNewFile();
                if (fname.endsWith("xlsx")) {
                   book = new XSSFWorkbook();
              //      book = WorkbookFactory.create(f);
                } else {
                    try {
                        throw new Exception("Debe ser un fichero Excel \"xlsx\"");
                    } catch (Exception ex) {
                        Logger.getLogger(CopySheets.class.getName()).log(Level.SEVERE, null, ex);
                    }
                }
            }
        

        for (int i= 0; i<files.length;i++){
            hoja[i]= new FileInputStream(files[i]);
        }
  //------------------Fin de Datos de Entrada a modificar -------------------------
        
    //    book = new XSSFWorkbook();
        ArrayList<FileInputStream> inList;
        inList = new ArrayList();
        for (int i=0; i<numero_de_ficheros; i++){
                   inList.add(hoja[i]);
          }
        System.out.println("Construido el Array de ficheros a importar");
//--------------------------------
        int i=0;
       for ( FileInputStream fin : inList) {
        Workbook b = new XSSFWorkbook(fin);
 //       Workbook b = WorkbookFactory.create(fin);
      //  for (int i = 0; i < b.getNumberOfSheets(); i++) {
        
            // not entering sheet name, because of duplicated names
            copySheets(book.createSheet(files[i].getName()),b.getSheetAt(numero_hoja));
            System.out.println("Importada la hoja "+numero_hoja+" del fichero "+files[i].getName());
            i++;
       }
//--------------------------------
       FileOutputStream fos = new FileOutputStream(f);
       book.write(fos);
       System.out.println("Se ha escrito el fichero "+fname);
       System.exit(0);
       
       
}


}
