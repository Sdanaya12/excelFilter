package com.ExcelFilter.automation.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class ExcelICBCLA {
    public static String getFilePath() {
        String filePath = "Resource/ICBCLA_Pruebas.xlsx";
        return filePath;
    }

    public static String getSheetName() {
        String sheetName = "ICBCLA";
        return sheetName;
    }

    public static int columnSuscriptor(){
        int columnSuscriptor = 1;
        return columnSuscriptor;
    }

    public static int columnClave(){
        int columnClave = 3;
        return columnClave;
    }

    public static int columnNombre(){
        int columnNombre = 4;
        return columnNombre;
    }

    public static int columnBloqueo(){
        int columnBloqueo = 9;
        return columnBloqueo;
    }

    public static int columnLastUtil(){
        int columnLastUtil = 23;
        return columnLastUtil;
    }

    public static int columnId(){
        int columnId = 28;
        return columnId;
    }

    public static int firstFilter() {
        int firtsFilter = 0;
        return firtsFilter;
    }

    public static String firstSeparation() {
        String firstSeparation = "E";
        return firstSeparation;
    }

    public static String secondSeparation() {
        String secondSeparation = ",";
        return secondSeparation;
    }

    public static String thirdSeparation() {
        String secondSeparation = ".";
        return secondSeparation;
    }

    public static String[][] excelFile_ICBCLA() throws IOException, InterruptedException {

        FileInputStream documento = new FileInputStream(getFilePath());
        XSSFWorkbook workbook = new XSSFWorkbook(documento);

        XSSFSheet sheet = workbook.getSheet(getSheetName());
        int contRow = sheet.getLastRowNum()-sheet.getFirstRowNum();
        int contColumn = sheet.getRow(0).getLastCellNum();
        int ii = 0;
        sheet.createRow(contRow + 1);
        Row row;
        Cell cell;
        String data;
        String [][] dataDinamica  = new String [contRow][contColumn];
        System.out.println("------------> EXCEL ORIGINAL: ICBCLA <------------");
        System.out.println("Data Original -> Filas: "+contRow+" ,Columnas: "+contColumn);
        for (int i = 0 ; i < contRow ; i++ ) {
            ii++;
            row = sheet.getRow(ii);
            for(int j = 0 ; j < contColumn ; j++) {
                cell = row.getCell(j);
                if(cell == null){
                    dataDinamica[i][j] = "";
                }else{
                    data = cell.toString();
                    dataDinamica[i][j] = data;
                }
            }
            //System.out.println(dataDinamica[i][0]+" || "+dataDinamica[i][1]+" || "+dataDinamica[i][2]);
        }
        return dataDinamica;
    }

    public static String[][] excelFilter_ICBCLA() throws IOException, InterruptedException {

        String [][] data = excelFile_ICBCLA();
        int firtsFilter = firstFilter();
        int contRow = data.length, contColumn = 6;
        int columnSuscriptor = columnSuscriptor(), columnClave = columnClave(), columnBloqueo = columnBloqueo(), contFilter = 0;
        int columnNombre = columnNombre(), columnLastUtil = columnLastUtil(), columnId = columnId();
        String [][] dataFilter = new String[contRow][contColumn];
        System.out.println("------------> EXCEL FILTRADO: ICBCLA <------------");
        for (int i = 0 ; i < contRow ; i++ ) {
            int dataBlock = Integer.parseInt(data[i][columnBloqueo]);
            if(dataBlock == firtsFilter ){
                   dataFilter[contFilter][0]=data[i][columnSuscriptor];
                   dataFilter[contFilter][1]=data[i][columnClave];
                   dataFilter[contFilter][2]=data[i][columnNombre];
                   dataFilter[contFilter][3]=data[i][columnBloqueo];
                   dataFilter[contFilter][4]=data[i][columnLastUtil];
                   dataFilter[contFilter][5]=data[i][columnId];
               contFilter ++;
            }
        }
        String [][] newDataFilter = new String [contFilter][contColumn];
        for (int i=0; i<contFilter;i++){
            for (int j=0; j<contColumn;j++){
                if(j==5){
                    double number = Double.valueOf(dataFilter[i][j]);
                    DecimalFormat decimalFormat = new DecimalFormat("#");
                    decimalFormat.setMaximumFractionDigits(10);
                    newDataFilter[i][j] = decimalFormat.format(number);
                }else{
                    newDataFilter[i][j]=dataFilter[i][j];
                }
            }
        }
        System.out.println("Data Filtrada -> Filas: "+contFilter+" ,Columnas: "+contColumn);
        return newDataFilter;
    }

    public static String[][] excelConvertDate_ICBCLA() throws IOException, InterruptedException {
        char charToAdd = '-';
        String[][] data = excelFilter_ICBCLA();
        int contRow = data.length, contColumn = data[0].length, columnLastUtil = 4;
        String date="";
        String [][] dataConvertDate = new String[contRow][contColumn];
        System.out.println("------------> EXCEL CONVERSIÓN FECHA: ICBCLA <------------");
        System.out.println("Data Filtrada -> Filas: "+contRow+" ,Columnas: "+contColumn);
        for (int i = 0 ; i < contRow ; i++ ) {
            for (int j = 0; j<contColumn;j++){
                String dataDate = data[i][columnLastUtil];
                String dateConvert = dataDate.substring(0,4)+charToAdd+dataDate.substring(4,6)+charToAdd+dataDate.substring(6);
                date = dateConvert;
            }
            data[i][columnLastUtil] = date;
        }
        saveExcel(contRow, data);
        return dataConvertDate;
    }

    public static void saveExcel(int contRow, String [][] dataSave) throws IOException, InterruptedException{
        int contData = 2;
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Datos Filtrados");
        Map<String, Object[]> datos = new TreeMap<String, Object[]>();

        datos.put("1", new Object[]{"Suscriptor", "Clave", "Nombre", "Bloqueo", "Ult Utilizacion", "ID"});
        for(int i=0;i<contRow;i++){
            String contDataString = String.valueOf(contData);
            datos.put(contDataString, new Object[]{dataSave[i][0], dataSave[i][1], dataSave[i][2], dataSave[i][3], dataSave[i][4], dataSave[i][5]});
            contData++;
        }
        Set keySet = datos.keySet();
        int numberRow = 0;
        for (Object key : keySet) {
            Row row = sheet.createRow(numberRow++);
            Object[] arrayObjects = datos.get(key);
            int numberCell = 0;
            for (Object obj : arrayObjects) {
                Cell cell = row.createCell(numberCell++);
                if (obj instanceof String) {
                    cell.setCellValue((String) obj);
                } else if (obj instanceof Integer) {
                    cell.setCellValue((Integer) obj);
                }
            }
        }
        FileOutputStream out = new FileOutputStream("Resource/ICBCLA_Pruebas_Filtrado.xlsx");
        workbook.write(out);
        out.close();
        System.out.println("------------> EXCEL CREADO: ICBCLA: ICBCLA_Pruebas_Filtrado <------------");
    }
}
