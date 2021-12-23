package com.ExcelFilter.automation.utils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.LocalDate;
import java.time.temporal.ChronoUnit;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class Excel {
    public static String getFilePath() {
        //String filePath = "Resource/DataPrueba.xlsx";
        String filePath = "Resource/Data.xlsx";
        return filePath;
    }

    public static String getSheetName() {
        String sheetName = "Data";
        return sheetName;
    }

    public static int columnUser(){
    int columnUser = 0;
    return columnUser;
    }

    public static int columnLogin() {
        int columnLogin = 1;
        return columnLogin;
    }

    public static int columnStatus() {
        int columnStatus = 2;
        return columnStatus;
    }

    public static int columnLastLogin(){
        int columnLastLogin = 5;
        return columnLastLogin;
    }

    public static String firstFilter() {
        String firtsFilter = "demo.datacredito.com.co";
        return firtsFilter;
    }

    public static String [] secondFilter() {
        String [] secondFilter = new String[2];
        secondFilter [0] = "ACTIVE";
        secondFilter [1] = "PASSWORD_EXPIRED";
        return secondFilter;
    }

    public static String [] thirdFilter() {
        String [] thirdFilter = new String[2];
        thirdFilter [0] = "2-";
        thirdFilter [1] = "88888888";
        return thirdFilter;
    }

    public static String fourthFilter(){
        String fourthFilter = "Usuario WS";
        return fourthFilter;
    }

    public static String firtsSeparation(){
        String firtsSeparation = "-";
        return firtsSeparation;
    }

    public static String secondSeparation(){
        String secondSeparation = "@";
        return secondSeparation;
    }

    public static String thirdSeparation(){
        String thirdSeparation = "T";
        return thirdSeparation;
    }

    public static int fourthSeparation(){
        int fourthSeparation = 120;
        return fourthSeparation;
    }

    public static String[][] excelFile() throws IOException {

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
        System.out.println("------------> EXCEL ORIGINAL <------------");
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

    public static String[][] excelFilter() throws IOException{

        String firtsFilter = firstFilter();
        String [] secondFilter = secondFilter();
        String [] thirdFilter = thirdFilter();
        String fourthFilter = fourthFilter();
        String [][] data = excelFile();
        int contRow = data.length, contColumn = data[0].length;
        int columnUser = columnUser(),columnLogin = columnLogin(), columnStatus = columnStatus(), contFilter = 0;
        String [][] dataFilter = new String[contRow][contColumn];
        System.out.println("------------> EXCEL FILTRADO <------------");
        for (int i = 0 ; i < contRow ; i++ ) {
            if(data[i][columnLogin].contains(firtsFilter)){
                if (data[i][columnStatus].contains(secondFilter[0]) || data[i][columnStatus].contains(secondFilter[1])){
                    if ((data[i][columnLogin].contains(thirdFilter[0]) || data[i][columnLogin].contains(thirdFilter[1])) == false){
                        if (data[i][columnLogin].contains(firtsSeparation())){
                            if (data[i][columnUser].contains(fourthFilter) != true){
                                //System.out.println(data[i][0]+" || "+data[i][j]+" || "+data[i][k]);
                                for (int l = 0; l<contColumn;l++){
                                    dataFilter[contFilter][l]=data[i][l];
                                }
                                contFilter ++;
                            }
                        }
                    }
                }
            }
        }
        System.out.println("Data Filtrada -> Filas: "+contFilter+" ,Columnas: "+contColumn);
        excelSeparation_Id(contFilter, contColumn, dataFilter);
        return dataFilter;
    }

    public static void excelSeparation_Id(int contRow, int contColumn, String [][] dataFilter) throws IOException {

        System.out.println("------------> EXCEL SEPARADO <------------");
        System.out.println("Data separada -> Filas: "+contRow+" ,Columnas: "+contColumn);
        String firtsSeparation = firtsSeparation(), secondSeparation = secondSeparation();
        int columnLogin = columnLogin();
        for (int i=0; i<contRow; i++){
            for (int j=0; j<contColumn; j++){
                if(columnLogin == j){
                    String [] idSeparation = dataFilter[i][j].split(firtsSeparation);
                    String [] id = idSeparation[1].split(secondSeparation);
                    //Imprime Usuario, indicador, documento, tipo, estado
                    //System.out.println(dataFilter[i][0]+" || "+idSeparation[0]+" || "+id[0]+" || "+id[1]+" || "+dataFilter[i][2]);
                    dataFilter[i][j] = id[0];
                }
            }
        }
        System.out.println("------------> EXCEL ORGANIZADO <------------");
        System.out.println("Data Organizada -> Filas: "+contRow+" ,Columnas: "+contColumn);
        for (int i=0; i<contRow; i++) {
            //System.out.println(dataFilter[i][0]+" || "+dataFilter[i][1]+" || "+dataFilter[i][2]);
        }
        excelSeparation_LastLogin(contRow, contColumn, dataFilter);
    }

    public static void excelSeparation_LastLogin(int contRow, int contColumn, String [][] dataFilter) throws IOException {
        int columnLastLogin = columnLastLogin(), contSeparation = 0;
        System.out.println("------------> EXCEL ORGANIZADO: LastLogin vacío <------------");
        String[][] dataSeparationLastLogin = new String[contRow][contColumn];
        String thirdSeparation = thirdSeparation();
        for (int i=0; i<contRow; i++){
            if(dataFilter[i][columnLastLogin] != ""){
                for (int j=0; j<contColumn; j++){
                    dataSeparationLastLogin[contSeparation][j]=dataFilter[i][j];
                }
                contSeparation++;
            }
        }
        System.out.println("Data LastLogin -> Filas: "+contSeparation+" ,Columnas: "+contColumn);

        for (int i=0; i<contSeparation; i++){
            for (int j=0; j<contColumn; j++) {
                if(columnLastLogin == j){
                    String [] dateSeparation = dataSeparationLastLogin[i][columnLastLogin].split(thirdSeparation);
                    dataSeparationLastLogin[i][j] = dateSeparation[0];
                }
                //System.out.println(dataSeparationLastLogin[i][j]);
            }
        }
        excelDateComparison(contSeparation, contColumn, dataSeparationLastLogin);
    }

    public static void excelDateComparison(int contRow, int contColumn, String [][] dataFilterSeparated ) throws IOException {
        int columnLastLogin = columnLastLogin(),fourthSeparation = fourthSeparation(), columnComparison = contColumn+1, contCandidates = 0;
        System.out.println("------------> EXCEL ORGANIZADO: Comparación de fechas <------------");
        String [][] dataComparison = new String[contRow][columnComparison];

        for (int i=0; i<contRow; i++){
            String  requestDate = dataFilterSeparated[i][columnLastLogin];
            LocalDate myDate = LocalDate.parse(requestDate);
            LocalDate currentDate = LocalDate.now();
            long numberOFDays = ChronoUnit.DAYS.between(myDate, currentDate);
            String numberDays = String.valueOf(numberOFDays);
            int totalDays = Integer.parseInt(numberDays);
            if(totalDays>fourthSeparation){
                for (int j=0; j<contColumn;j++){
                    //System.out.println(myDate+" VS "+currentDate+" = "+numberDays);
                    dataComparison[contCandidates][j]=dataFilterSeparated[i][j];
                    dataComparison[contCandidates][contColumn] = numberDays;
                }
                contCandidates++;
            }
        }
        saveExcel(contCandidates, dataComparison);
        System.out.println("Data Comparison -> Filas: "+contCandidates+" ,Columnas: "+columnComparison);
    }

    public static void saveExcel(int contRow, String [][] dataSave) throws IOException {
        int contData = 2;
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Datos Filtrados");
        Map<String, Object[]> datos = new TreeMap<String, Object[]>();

        datos.put("1", new Object[]{"User", "Login", "Status", "Activation Date", "Authentication Source", "Last Login", "Last Password Change", "Days Inactive"});
        for(int i=0;i<contRow;i++){
            String contDataString = String.valueOf(contData);
            datos.put(contDataString, new Object[]{dataSave[i][0], dataSave[i][1], dataSave[i][2], dataSave[i][3], dataSave[i][4], dataSave[i][5], dataSave[i][6], dataSave[i][7]});
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
        FileOutputStream out = new FileOutputStream("Resource/DataFinal.xlsx");
        workbook.write(out);
        out.close();
    }
}
