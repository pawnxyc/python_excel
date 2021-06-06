package com;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class DealExcel {

    private static XSSFWorkbook hssfSheet = null;
    private static XSSFWorkbook target = null;
    private static HashMap<String,Integer> map  = new HashMap<>();
    private static HashMap<Integer,Integer> sheetMapCount  = new HashMap<>();


    public static void main(String[] args) throws IOException, InvalidFormatException {

        String fileloc = "/Users/xiapawn/IdeaProjects/python_excel/src/main/resources/excel/西昌·邦泰花园城3期.xlsx";
        String targetloc = "/Users/xiapawn/IdeaProjects/python_excel/src/main/resources/excel/target.xlsx";
        hssfSheet = new XSSFWorkbook(new File(fileloc));
        initTarget();

        hssfSheet.forEach(sheetOne ->{

            for (int rownum = 0; rownum <= sheetOne.getLastRowNum(); rownum++) {


                //获取到每一行
                XSSFRow sheetRow = (XSSFRow)sheetOne.getRow(rownum);

                String value = getValue(sheetRow.getCell(1));
                Pattern p = Pattern.compile("[0-9]{3,}");
                Matcher matcher = p.matcher(value);

                if (sheetRow == null) {
                    continue;
                }
                if (getValue(sheetRow.getCell(1)).length()<5){
                    continue;
                }
//                如果是包含在筛选条件中 则给指定的value 的sheet来添加行和列 并赋予内容
                if (matcher.find()||map.containsKey(getValue(sheetRow.getCell(1)).substring(0,4))){
                    Integer integer = map.get(getValue(sheetRow.getCell(1)).substring(0, 4));
                    integer = integer==null?9:integer;
                    Integer sheetRowCount = sheetMapCount.get(integer);
                    XSSFSheet targetSheetAt = target.getSheetAt(integer);
                    sheetMapCount.put(integer,sheetRowCount+1);
                    int lastRowNum = sheetRowCount;
                    XSSFRow row = targetSheetAt.createRow(lastRowNum ==0?0:lastRowNum);
                    for (int cellnum = 0; cellnum <= sheetRow.getLastCellNum(); cellnum++) {
                        XSSFCell cell = sheetRow.getCell(cellnum);
                        if (cell == null) {
                            continue;
                        }

                        XSSFCell targetcell = row.createCell(cellnum, cell.getCellType());
                            if (cell.getCellType() == CellType.BOOLEAN) {
                                targetcell.setCellValue(cell.getBooleanCellValue());

                            }else if (cell.getCellType() == CellType.NUMERIC) {
                                targetcell.setCellValue(cell.getNumericCellValue());

                            }else if(cell.getCellType() == CellType.STRING){
                                targetcell.setCellValue(cell.getStringCellValue());

                            }else if(cell.getCellType() == CellType.FORMULA){
                                targetcell.setCellValue(cell.getCellFormula());

                            }else if (cell.getCellType() == CellType.BLANK){
                                targetcell.setCellValue("未知类型");

                            }else if(cell.getCellType() == CellType.ERROR){
                                targetcell.setCellValue("未知类型");
                            }else{
                                targetcell.setCellValue("未知类型");
                            }

                    }
                continue;
            }


            }
    });
        target.write(new FileOutputStream(targetloc));

    }

    private static String getValue(XSSFCell hssfCell) {
        String cellValue = "";
        if(hssfCell == null){
            return cellValue;
        }
        //把数字当成String来读，避免出现1读成1.0的情况
        hssfCell.setCellType(CellType.STRING);

        //hssfCell.getCellType() 获取当前列的类型
        if (hssfCell.getCellType() == CellType.BOOLEAN) {
            cellValue = String.valueOf(hssfCell.getBooleanCellValue());
        }else if (hssfCell.getCellType() == CellType.NUMERIC) {
            cellValue =  String.valueOf(hssfCell.getNumericCellValue());
        }else if(hssfCell.getCellType() == CellType.STRING){
            cellValue =  String.valueOf(hssfCell.getStringCellValue());
        }else if(hssfCell.getCellType() == CellType.FORMULA){
            cellValue =  String.valueOf(hssfCell.getCellFormula());
        }else if (hssfCell.getCellType() == CellType.BLANK){
            cellValue = " ";
        }else if(hssfCell.getCellType() == CellType.ERROR){
            cellValue = "非法字符";
        }else{
            cellValue = "未知类型";
        }
        return cellValue;
    }

    static void initTarget(){
        map.put("0105",0);
        map.put("0104",1);
        map.put("0101",1);
        map.put("0106",1);
        map.put("0109",2);
        map.put("0110",2);
        map.put("0111",3);
        map.put("0112",4);
        map.put("0113",5);
        map.put("0114",6);
        map.put("0115",7);
        map.put("0117",8);

        sheetMapCount.put(0,0);
        sheetMapCount.put(1,0);
        sheetMapCount.put(2,0);
        sheetMapCount.put(3,0);
        sheetMapCount.put(4,0);
        sheetMapCount.put(5,0);
        sheetMapCount.put(6,0);
        sheetMapCount.put(7,0);
        sheetMapCount.put(8,0);
        sheetMapCount.put(9,0);

        String[] allKinds = {"钢筋+砼","砌体","屋面","楼地面","墙面","天棚","油漆","线条","模板","其他"};
        target = new XSSFWorkbook();
        for (String one:allKinds) {
            target.createSheet(one);
        }

    }
}
