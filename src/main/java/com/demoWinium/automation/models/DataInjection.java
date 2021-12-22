package com.demoWinium.automation.models;

import com.demoWinium.automation.utils.Excel;

import java.io.IOException;

public class DataInjection {
    public String getFilePath() {
        String filePath = "Resource/DataPrueba.xlsx";
        return filePath;
    }

    public String getSheetName() {
        String sheetName = "Data";
        return sheetName;
    }

    //Acá realizar métodos de filtrado
    public void data(int contFilas, int contColumnas) throws IOException {
        System.out.println("Filas: "+contFilas+", Columnas: "+contColumnas);
        String[][] data = Excel.excelFile();
        String[] infoCell = new String[contColumnas];
        for (int i=0; i<contFilas; i++ ){
            for (int j=0; j<contColumnas;j++){
                infoCell[j] = data[i][j];
            }
        }
    }
}
