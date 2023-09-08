package model;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XlsxContainer {
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private String nameFile;
    private int cadRow;
    private int costRow;
    private int areaRow;
    private int indexSheetRow;

    public int getIndexSheetRow() {
        return indexSheetRow;
    }

    public void setIndexSheetRow(int indexSheetRow) {
        this.indexSheetRow = indexSheetRow;
    }

    public int getCadRow() {
        return cadRow;
    }

    public void setCadRow(int cadRow) {
        this.cadRow = cadRow;
    }

    public int getCostRow() {
        return costRow;
    }

    public void setCostRow(int costRow) {
        this.costRow = costRow;
    }

    public int getAreaRow() {
        return areaRow;
    }

    public void setAreaRow(int areaRow) {
        this.areaRow = areaRow;
    }

    public String getNameFile() {
        return nameFile;
    }

    public void setNameFile(String nameFile) {
        this.nameFile = nameFile;
    }

    public XSSFWorkbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(XSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    public XSSFSheet getSheet() {
        return sheet;
    }

    public void setSheet(XSSFSheet sheet) {
        this.sheet = sheet;
    }
}
