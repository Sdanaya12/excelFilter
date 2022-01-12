package com.ExcelFilter.automation.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.temporal.ChronoUnit;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class ExcelDataCrossover {

    public static int columnKeys_ICBCLA(){
        int columnKeys_ICBCLA = 1;
        return columnKeys_ICBCLA;
    }

    public static int columnLastUtil_ICBCLA(){
        int columnLastUtil_ICBCLA = 4;
        return columnLastUtil_ICBCLA;
    }

    public static int columnId_ICBCLA(){
        int columnId_ICBCLA = 5;
        return columnId_ICBCLA;
    }

    public static int columnComparison_ICBCLA(){
        int columnComparison_ICBCLA = 6;
        return columnComparison_ICBCLA;
    }

    public static int columnTotalDays_ICBCLA(){
        int columnTotalDays_ICBCLA = 7;
        return columnTotalDays_ICBCLA;
    }

    public static int columnUser_Okta(){
        int columnUser_Okta = 0;
        return columnUser_Okta;
    }

    public static int columnId_Okta(){
        int columnId_Okta = 1;
        return columnId_Okta;
    }

    public static int columnLogin_Okta(){
        int columnLogin_Okta = 3;
        return columnLogin_Okta;
    }

    public static int columnAuthentication_Okta(){
        int columnAuthentication_Okta = 4;
        return columnAuthentication_Okta;
    }

    public static int columnLastUtil_Okta(){
        int columnLastUtil_Okta = 5;
        return columnLastUtil_Okta;
    }

    public static int firstSeparation(){
        int firstSeparation = 120;
        return firstSeparation;
    }

    public static String secondSeparation(){
        String secondSeparation = "OK";
        return secondSeparation;
    }

    public static String thirdSeparation(){
        String thirdSeparation = "1";
        return thirdSeparation;
    }

    public static String firstIndicator(){
        String firstIndicator = "Está presente";
        return firstIndicator;
    }

    public static String[][] excelComparison() throws IOException{
        String[][] dataOkta = ExcelOkta.excelDateComparison();
        String[][] dataICBCLA = ExcelICBCLA.excelConvertDate_ICBCLA();
        String firstIndicator = firstIndicator();
        int contRow_Okta = dataOkta.length, contColumn_Okta = dataOkta[0].length, columnId_Okta = columnId_Okta(), columnAuthentication_Okta = columnAuthentication_Okta();
        int contRow_ICBCLA = dataICBCLA.length, contColumn_ICBCLA = dataICBCLA[0].length, columnId_ICBCLA = columnId_ICBCLA();
        int contRowCoincidences_ICBCLA = 0, contRowCoincidences_Okta = 0, ii=0, jj=0, ll=0;
        String [][] dataICBCLA_Comparison = new String[contRow_ICBCLA][contColumn_ICBCLA+1];
        System.out.println("Matriz Okta: "+contRow_Okta+" x "+contColumn_Okta+" ,Matriz ICBCLA: "+contRow_ICBCLA+" x " +contColumn_ICBCLA);
        for (int i = 0; i<contRow_Okta; i++){
            for (int j = 0; j<contRow_ICBCLA; j++){
                if(dataOkta[i][columnId_Okta].equals(dataICBCLA[j][columnId_ICBCLA])){//Comparación de Okta vs ICBCLA
                    dataOkta[i][columnAuthentication_Okta] = firstIndicator;
                    for (int k = 0; k<contColumn_ICBCLA+1; k++){
                        if(k==contColumn_ICBCLA){
                            dataICBCLA_Comparison[j][k] = firstIndicator;
                        }else{
                            dataICBCLA_Comparison[j][k] = dataICBCLA[j][k];
                        }
                    }
                    contRowCoincidences_ICBCLA++;
                }else{
                    for (int k = 0; k<contColumn_ICBCLA; k++){
                        dataICBCLA_Comparison[j][k] = dataICBCLA[j][k];
                    }
                }
                //System.out.println("Id: "+dataICBCLA_Comparison[j][columnId_ICBCLA]+", Comparación: "+dataICBCLA_Comparison[j][contColumn_ICBCLA]);
            }
        }
        for(int i = 0; i<contRow_Okta; i++){
            if(dataOkta[i][columnAuthentication_Okta] != firstIndicator){
                contRowCoincidences_Okta++;
            }
        }
        int newContRowCoincidences_Okta = contRow_Okta-contRowCoincidences_Okta;
        String [][] dataComparisonFilter_ICBCLA = new String[contRowCoincidences_ICBCLA][contColumn_ICBCLA+1];
        String [][] dataComparisonFilter_Okta = new String[contRowCoincidences_Okta][contColumn_Okta];
        String [][] dataComparisonFilterPresence_Okta = new String[newContRowCoincidences_Okta][contColumn_Okta];
        for (int i = 0; i<contRow_ICBCLA; i++){
            if(dataICBCLA_Comparison[i][contColumn_ICBCLA] != null){
                for (int j = 0; j<contColumn_ICBCLA+1; j++){
                    dataComparisonFilter_ICBCLA[ii][j] = dataICBCLA_Comparison[i][j];
                }
                ii++;
            }
        }

        for (int l = 0; l<contRow_Okta; l++){
            if(dataOkta[l][columnAuthentication_Okta] != firstIndicator){
                for (int m = 0; m<contColumn_Okta; m++){
                    dataComparisonFilter_Okta[jj][m] = dataOkta[l][m];
                }
                jj++;
            }
        }

        for (int l = 0; l<contRow_Okta; l++){
            if(dataOkta[l][columnAuthentication_Okta] == firstIndicator){
                for (int m = 0; m<contColumn_Okta; m++){
                    dataComparisonFilterPresence_Okta[ll][m] = dataOkta[l][m];
                }
                ll++;
            }
        }

        saveExcel_Okta(contRowCoincidences_Okta, dataComparisonFilter_Okta);
        saveExcel_ICBCLA(contRowCoincidences_ICBCLA, dataComparisonFilter_ICBCLA);
        excelVerifyKeys_ICBCLA(dataComparisonFilter_ICBCLA, dataComparisonFilterPresence_Okta);
        return dataComparisonFilter_ICBCLA;
    }

    public static String[][] excelVerifyKeys_ICBCLA(String [][] dataComparisonFilter, String [][] dataComparisonFilterPresence_Okta ) throws IOException {
        String secondSeparation = secondSeparation(), thirdSeparation = thirdSeparation();
        int contRow = dataComparisonFilter.length, contColumn = dataComparisonFilter[0].length;
        int contRowPresenceOkta = dataComparisonFilterPresence_Okta.length, contColumnPresenceOkta = dataComparisonFilterPresence_Okta[0].length, columnId_Okta = columnId_Okta();
        int columnKeys_ICBCLA = columnKeys_ICBCLA(), columnId_ICBCLA = columnId_ICBCLA(), columnLastUtil_ICBCLA = columnLastUtil_ICBCLA(), firstSeparation = firstSeparation();
        int contComparison = 0, columnLastUtil_Okta = columnLastUtil_Okta(), columnUser_Okta = columnUser_Okta(), columnLogin_Okta=columnLogin_Okta(), ii = 0, k=0, contId = 0;;
        int contSmallerThan = 0, columnComparison_ICBCLA = columnComparison_ICBCLA(), columnTotalDays_ICBCLA = columnTotalDays_ICBCLA();
        String[][] dataComparisonLastUtil = new String[contRow][contColumn+1];
        String[][] dataComparisonSmallerThan = new String[contRow][contColumn+1];
        String[][] dataFinish = new String[contRow][contColumn+1];
        String[][] reportFinish = new String[contRowPresenceOkta][contColumnPresenceOkta-1];
        for (int i=0; i<contRow; i++){
            String  requestDate = dataComparisonFilter[i][columnLastUtil_ICBCLA];
            LocalDate myDate = LocalDate.parse(requestDate);
            LocalDate currentDate = LocalDate.now();
            long numberOFDays = ChronoUnit.DAYS.between(myDate, currentDate);
            String numberDays = String.valueOf(numberOFDays);
            for (int j=0; j<contColumn+1;j++){
                if(j==contColumn){
                    dataComparisonLastUtil[contComparison][j] = numberDays;
                }else{
                    dataComparisonLastUtil[contComparison][j] = dataComparisonFilter[i][j];
                }
            }
            contComparison++;
        }
        for (int i=0; i<contRow; i++){
            int totalDays = Integer.parseInt(dataComparisonLastUtil[i][columnTotalDays_ICBCLA]);
            if (totalDays<firstSeparation){
                for(int j=0; j<contColumn; j++){
                    dataComparisonSmallerThan[contSmallerThan][j]=dataComparisonLastUtil[i][j];
                }
                contSmallerThan++;
            }
        }
        for (int i=0; i<contSmallerThan; i++){
            String Id = dataComparisonSmallerThan[i][columnId_ICBCLA];
            for (int j = 0; j<contRow; j++){
                if(Id.equals(dataComparisonLastUtil[j][columnId_ICBCLA])){
                    dataComparisonLastUtil[j][columnComparison_ICBCLA] = secondSeparation;
                }
            }
        }
        for (int i=0; i<contRow; i++){
            String comparison = dataComparisonLastUtil[i][columnComparison_ICBCLA];
            String Id = dataComparisonLastUtil[i][columnId_ICBCLA];
            if(comparison.equals(secondSeparation) != true){
                for(int j=0; j<contRow; j++){
                    if(Id.equals(dataComparisonLastUtil[j][columnId_ICBCLA])){
                        contId++;
                        String cont_Id = String.valueOf(contId);
                        dataComparisonLastUtil[j][columnComparison_ICBCLA] = cont_Id;
                    }
                }
                contId = 0;
            }
        }
        for (int i=0; i<contRow; i++){
            String indicator = dataComparisonLastUtil[i][columnComparison_ICBCLA];
            if(indicator.equals(thirdSeparation)){
                for (int j=0; j<contColumn+1; j++){
                    dataFinish[ii][j] = dataComparisonLastUtil[i][j];
                }
                ii++;
            }
        }
        int newContRow = ii;
        for(int i=0; i<contRowPresenceOkta;i++){
            for (int j=0; j<newContRow; j++){
                if(dataComparisonFilterPresence_Okta[i][columnId_Okta].equals(dataFinish[j][columnId_ICBCLA])){
                    reportFinish[k][0]=dataFinish[j][columnId_ICBCLA];
                    reportFinish[k][1]=dataFinish[j][columnKeys_ICBCLA];
                    reportFinish[k][2]=dataFinish[j][columnLastUtil_ICBCLA];
                    reportFinish[k][3]=dataFinish[j][columnTotalDays_ICBCLA];
                    reportFinish[k][4]=dataComparisonFilterPresence_Okta[i][columnUser_Okta];
                    reportFinish[k][5]=dataComparisonFilterPresence_Okta[i][columnLogin_Okta];
                    reportFinish[k][6]=dataComparisonFilterPresence_Okta[i][columnLastUtil_Okta];
                    k++;
                }
            }
        }
        saveExcel(k, reportFinish);
        return dataFinish;
    }

    public static void saveExcel_Okta(int contRow, String [][] dataSave) throws IOException{
        int contData = 2;
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("No presentes en ICBCLA");
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
        FileOutputStream out = new FileOutputStream("Resource/ListadoTotalOktaAP_ELIMINAR.xlsx");
        workbook.write(out);
        out.close();
        System.out.println("------------> EXCEL CREADO: Okta: ListadoTotalOktaAP_ELIMINAR <------------");
    }

    public static void saveExcel_ICBCLA(int contRow, String [][] dataSave) throws IOException{
        int contData = 2;
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Datos Filtrados");
        Map<String, Object[]> datos = new TreeMap<String, Object[]>();

        datos.put("1", new Object[]{"Suscriptor", "Clave", "Nombre", "Bloqueo", "Ult Utilizacion", "ID", "Comparación"});
        for(int i=0;i<contRow;i++){
            String contDataString = String.valueOf(contData);
            datos.put(contDataString, new Object[]{dataSave[i][0], dataSave[i][1], dataSave[i][2], dataSave[i][3], dataSave[i][4], dataSave[i][5], dataSave[i][6]});
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
        FileOutputStream out = new FileOutputStream("Resource/ICBCLA_Pruebas_Coincidencias.xlsx");
        workbook.write(out);
        out.close();
        System.out.println("------------> EXCEL CREADO: ICBCLA: ICBCLA_Pruebas_Coincidencias <------------");
    }

    public static void saveExcel(int contRow, String [][] dataSave) throws IOException{
        int contData = 2;
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Datos Filtrados");
        Map<String, Object[]> datos = new TreeMap<String, Object[]>();

        datos.put("1", new Object[]{"Id", "Clave", "Ult. Util ICBCLA", "Días inactivo", "Usuario", "Login", "Ult Utilizacion"});
        for(int i=0;i<contRow;i++){
            String contDataString = String.valueOf(contData);
            datos.put(contDataString, new Object[]{dataSave[i][0], dataSave[i][1], dataSave[i][2], dataSave[i][3], dataSave[i][4], dataSave[i][5], dataSave[i][6]});
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
        FileOutputStream out = new FileOutputStream("Resource/ReporteFinal.xlsx");
        workbook.write(out);
        out.close();
        System.out.println("------------> EXCEL CREADO: ReporteFinal <------------");
    }
}
