package com.ExcelFilter.automation.utils;
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

public class ExcelOkta {

    public static String getFilePath() {
        String filePath = "Resource/ListadoTotalOktaAP.xlsx";
        return filePath;
    }

    public static String getSheetName() {
        String sheetName = "PasswordHealthReport_phr12iqt37";
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

    public static int columnActivation(){
        int columnActivation = 3;
        return columnActivation;
    }

    public static int columnLastLogin(){
        int columnLastLogin = 5;
        return columnLastLogin;
    }

    public static String firstFilter() {
        String firtsFilter = "demo.";
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
        thirdFilter [1] = "88888";
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

    public static String[][] excelFile_Okta() throws IOException {

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
        //System.out.println("------------> EXCEL ORIGINAL: Okta <------------");
        //System.out.println("Data Original -> Filas: "+contRow+" ,Columnas: "+contColumn);
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

    public static String[][] excelFilter_Okta() throws IOException{

        String firtsFilter = firstFilter();
        String [] secondFilter = secondFilter();
        String [] thirdFilter = thirdFilter();
        String fourthFilter = fourthFilter();
        String [][] data = excelFile_Okta();
        int contRow = data.length, contColumn = data[0].length;
        int columnUser = columnUser(),columnLogin = columnLogin(), columnStatus = columnStatus(), contFilter = 0;
        String [][] dataFilter = new String[contRow][contColumn];
        //System.out.println("------------> EXCEL FILTRADO: Okta <------------");
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
        //System.out.println("Data Filtrada -> Filas: "+contFilter+" ,Columnas: "+contColumn);
        String [][] newDataFilter = new String[contFilter][contColumn];
        for (int i =0; i<contFilter;i++){
            for (int j=0; j<contColumn; j++){
                newDataFilter[i][j]=dataFilter[i][j];
            }
        }
        return newDataFilter;
    }

    public static String [][] excelSeparation_Id() throws IOException {

        String [][] dataFilter = excelFilter_Okta();
        int contRow = dataFilter.length, contColumn = dataFilter[0].length;
        int columnActivation=columnActivation();
        //System.out.println("------------> EXCEL SEPARADO: Okta <------------");
        //System.out.println("Data separada -> Filas: "+contRow+" ,Columnas: "+contColumn);
        String firtsSeparation = firtsSeparation(), secondSeparation = secondSeparation();
        int columnLogin = columnLogin(), cont = 0;
        for (int i=0; i<contRow; i++){
            for (int j=0; j<contColumn; j++){
                if(columnLogin == j){
                    dataFilter[i][columnActivation] = dataFilter[i][j];
                    String [] idSeparation = dataFilter[i][j].split(firtsSeparation);
                    String [] id = idSeparation[1].split(secondSeparation);
                    //Imprime Usuario, indicador, documento, tipo, estado
                    //System.out.println(dataFilter[i][0]+" || "+idSeparation[0]+" || "+id[0]+" || "+id[1]+" || "+dataFilter[i][2]);
                    dataFilter[i][j] = id[0];
                }
            }
        }
        //System.out.println("------------> EXCEL ORGANIZADO: Okta <------------");
        //System.out.println("Data Organizada -> Filas: "+contRow+" ,Columnas: "+contColumn);
        return dataFilter;
    }

    public static String[][] excelSeparation_LastLogin() throws IOException {
        int columnLastLogin = columnLastLogin(), contSeparation = 0;
        String [][] dataFilter = excelSeparation_Id();
        int contRow = dataFilter.length, contColumn = dataFilter[0].length;
        //System.out.println("------------> EXCEL ORGANIZADO: Okta: LastLogin vac??o <------------");
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
        //System.out.println("Data LastLogin -> Filas: "+contSeparation+" ,Columnas: "+contColumn);
        String [][] newDataSeparationLastLogin = new String[contSeparation][contColumn];
        for (int i=0; i<contSeparation; i++){
            for (int j=0; j<contColumn; j++) {
                if(columnLastLogin == j){
                    String [] dateSeparation = dataSeparationLastLogin[i][columnLastLogin].split(thirdSeparation);
                    dataSeparationLastLogin[i][j] = dateSeparation[0];

                }
                //System.out.println(dataSeparationLastLogin[i][j]);
            }
        }

        for (int i =0; i<contSeparation;i++){
            for (int j=0; j<contColumn; j++){
                newDataSeparationLastLogin[i][j] = dataSeparationLastLogin[i][j];
            }
        }
        return newDataSeparationLastLogin;
    }

    public static String[][] excelDateComparison() throws IOException {
        String [][] dataFilterSeparated = excelSeparation_LastLogin();
        int contRow = dataFilterSeparated.length, contColumn = dataFilterSeparated[0].length;
        int columnLastLogin = columnLastLogin(),fourthSeparation = fourthSeparation(), columnComparison = contColumn+1, contCandidates = 0;
        //System.out.println("------------> EXCEL ORGANIZADO: Okta: Comparaci??n de fechas <------------");
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
        String [][] newDataComparison = new String[contCandidates][columnComparison];
        for(int i=0;i<contCandidates;i++){
            for (int j=0; j<columnComparison; j++){
                newDataComparison[i][j] = dataComparison[i][j];
            }
        }
        saveExcel(contCandidates, newDataComparison);
        //System.out.println("Data Comparison -> Filas: "+contCandidates+" ,Columnas: "+columnComparison);
        return newDataComparison;
    }

    public static void saveExcel(int contRow, String [][] dataSave) throws IOException{
        int contData = 2;
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Datos Filtrados");
        Map<String, Object[]> datos = new TreeMap<String, Object[]>();

        datos.put("1", new Object[]{"User", "Id", "Status", "Login", "Authentication Source", "Last Login", "Last Password Change", "Days Inactive"});
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
        FileOutputStream out = new FileOutputStream("Resource/ListadoTotalOktaAP_Filtrado.xlsx");
        workbook.write(out);
        out.close();
        System.out.println("------------> EXCEL CREADO: Okta: ListadoTotalOktaAP_Filtrado <------------");
    }
}
