package com.org.combineexel.util;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.*;

public class MergeMultipleXlsFilesInDifferentSheet{
    public static void mergeExcelFiles() throws IOException {

        XSSFWorkbook outputWorkBook = new XSSFWorkbook();
        File outputFile = new File("D:\\exel\\3.xlxs");

        List<String> fileNames = new ArrayList<String>(){
            {
                add("D:\\exel\\1.xlsx");
                add("D:\\exel\\2.xlsx");
            }
        };
        XSSFSheet sheet = outputWorkBook.createSheet("consolidated sheet");
        for (String fileName : fileNames){
            File file = new File(fileName);
            if (file.isFile()){
                FileInputStream fin = new FileInputStream(file);
                XSSFWorkbook b = new XSSFWorkbook(fin);
                for (int i = 0; i < b.getNumberOfSheets(); i++) {
                    for(int worksheetIndex = 0; worksheetIndex<b.getNumberOfSheets(); worksheetIndex++)
                    {
                        String worksheetName = b.getSheetName(worksheetIndex);
                        System.out.println("test " + worksheetName);
                        copySheets(sheet, b.getSheetAt(i));
                        System.out.println("Copying..");
                    }
                }
            }
        }
        try {
            writeFile(outputWorkBook, outputFile);
        }catch(Exception e) {
            e.printStackTrace();
        }

    }

    protected static void writeFile(XSSFWorkbook book, File file) throws Exception {

    try{
        FileOutputStream out = new FileOutputStream(file);
        book.write(out);
        out.close();
        System.out.println(file+ " is written successfully..");
    } catch (FileNotFoundException e) {
        e.printStackTrace();
    } catch (IOException e) {
        e.printStackTrace();
    }
}


    private static void copySheets(XSSFSheet newSheet, XSSFSheet sheet){

        copySheets( newSheet, sheet, true);
    }

    private static void copySheets( XSSFSheet newSheet, XSSFSheet sheet, boolean copyStyle){
        int newRownumber = newSheet.getLastRowNum()  + 1;
        int maxColumnNum = 0;
        Map<Integer, XSSFCellStyle> styleMap = (copyStyle) ? new HashMap<Integer, XSSFCellStyle>() : null;

        for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
            XSSFRow srcRow = sheet.getRow(i);
            XSSFRow destRow = newSheet.createRow(i + newRownumber);
            if (srcRow != null) {
                //copyRow(newWorkbook, sheet, newSheet, srcRow, destRow, styleMap);
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


    public static void copyRow( XSSFSheet srcSheet, XSSFSheet destSheet, XSSFRow srcRow, XSSFRow destRow, Map<Integer, XSSFCellStyle> styleMap) {
        destRow.setHeight(srcRow.getHeight());
        Set<CellRangeAddress> mergedRegions = new TreeSet<CellRangeAddress>();


        int deltaRows = destRow.getRowNum()-srcRow.getRowNum();


        int j = srcRow.getFirstCellNum();
        if(j<0){j=0;}
        for (; j <= srcRow.getLastCellNum(); j++) {

            XSSFCell oldCell = srcRow.getCell(j);
            XSSFCell newCell = destRow.getCell(j);
            if (oldCell != null) {
                if (newCell == null) {
                    newCell = destRow.createCell(j);
                }

                copyCell( oldCell, newCell, styleMap);CellRangeAddress mergedRegion = getMergedRegion( srcSheet, srcRow.getRowNum(), (short)oldCell.getColumnIndex());
                if (mergedRegion != null) {

                    CellRangeAddress newMergedRegion = new CellRangeAddress(mergedRegion.getFirstRow()+deltaRows, mergedRegion.getLastRow()+deltaRows, mergedRegion.getFirstColumn(),  mergedRegion.getLastColumn());
                    //System.out.println("New merged region: " + newMergedRegion.toString());

                    if (isNewMergedRegion( newMergedRegion, mergedRegions)) {
                        mergedRegions.add(newMergedRegion);
                        destSheet.addMergedRegion(newMergedRegion);
                    }
                }


            }
        }
    }


    public static void copyCell( XSSFCell oldCell, XSSFCell newCell, Map<Integer, XSSFCellStyle> styleMap) {
        if(styleMap != null) {
            int stHashCode = oldCell.getCellStyle().hashCode();
            XSSFCellStyle newCellStyle = styleMap.get(stHashCode);
            if(newCellStyle == null){
                //newCellStyle = newWorkbook.createCellStyle();
                newCellStyle = newCell.getSheet().getWorkbook().createCellStyle();
                newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
                styleMap.put(stHashCode, newCellStyle);
            }
            newCell.setCellStyle(newCellStyle);
        }
        switch(oldCell.getCellType()) {
            case XSSFCell.CELL_TYPE_STRING:
                newCell.setCellValue(oldCell.getRichStringCellValue());
                break;
            case XSSFCell.CELL_TYPE_NUMERIC:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            case XSSFCell.CELL_TYPE_BLANK:
                newCell.setCellType(XSSFCell.CELL_TYPE_BLANK);
                break;
            case XSSFCell.CELL_TYPE_BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case XSSFCell.CELL_TYPE_ERROR:
                newCell.setCellErrorValue(oldCell.getErrorCellValue());
                break;
            case XSSFCell.CELL_TYPE_FORMULA:
                newCell.setCellFormula(oldCell.getCellFormula());
                break;
            default:
                break;
        }
    }


    public static CellRangeAddress getMergedRegion( XSSFSheet sheet, int rowNum, short cellNum) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress merged = (CellRangeAddress) sheet.getMergedRegion(i);
            if (merged.isInRange(rowNum, cellNum)) {
                return merged;
            }
        }
        return null;
    }




    private static boolean isNewMergedRegion(CellRangeAddress newMergedRegion, Collection<CellRangeAddress> mergedRegions) {
        return !mergedRegions.contains(newMergedRegion);
    }


    public static void main(String[] args) {
        try {
            mergeExcelFiles();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}