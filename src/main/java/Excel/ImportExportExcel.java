package Excel;


import model.CadNumber;
import model.XlsxContainer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

public class ImportExportExcel {



    public Set<String> getSetCad(XlsxContainer container) throws IOException, InvalidFormatException {

        XSSFSheet sheet = container.getSheet();

        Set<String> cadnumSet = new LinkedHashSet<>();
        Iterator<Row> iteratorRows =sheet.rowIterator();

        if(iteratorRows.hasNext()){
            iteratorRows.next();
        }
        XSSFCell cellCost = null;
        XSSFCell cellNumber = null;

        while (iteratorRows.hasNext()) {
            XSSFRow row = (XSSFRow) iteratorRows.next();
            cellCost = row.getCell(container.getCostRow()-1);
            cellNumber = row.getCell(container.getCadRow()-1);

            if (cellCost != null) {
                if (cellNumber.getStringCellValue().matches("(\\d{2}):(\\d{2}):(\\d{7}):(\\d+)")) {
                    cadnumSet.add(cellNumber.getStringCellValue());
                }
            }
        }
        return cadnumSet;
    }

    public void updateCostToExcel(XlsxContainer xlsxContainer, List<CadNumber> cadNumList) throws InvalidFormatException, IOException {

        XSSFWorkbook workbook = xlsxContainer.getWorkbook();
        XSSFSheet sheet = workbook.getSheetAt(xlsxContainer.getIndexSheetRow()-1);
        Iterator<Row> iteratorRows = sheet.rowIterator();

        XSSFCell cellCost = null;
        XSSFCell cellNumber = null;

        while (iteratorRows.hasNext()){
            XSSFRow row = (XSSFRow) iteratorRows.next();
            cellCost = row.getCell(xlsxContainer.getCostRow()-1);
            cellNumber = row.getCell(xlsxContainer.getCadRow()-1);


            for (int i = 0; i < cadNumList.size(); i++) {
                if ( cellNumber != null && cadNumList.get(i).getCadNumber().equals(cellNumber.getStringCellValue())) {
                    cellCost.setCellValue(cadNumList.get(i).getCost());

                }
            }
        }
        String pattern = "dd-MM-YYYY";
        Date date = new Date();
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat(pattern);
        File file = new File("C:/конвертированные выписки/"+" обновлен от "+simpleDateFormat.format(date) +" "+ xlsxContainer.getNameFile()+".xlsx");
        file.getParentFile().mkdirs();

        try (FileOutputStream outFile = new FileOutputStream(file)) {
            workbook.write(outFile);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public XlsxContainer getXSSFSheetAndWorkbook(File fileExcel, int index) throws  InvalidFormatException, NotOfficeXmlFileException,IllegalArgumentException,IOException {
        XlsxContainer xlsxContainer = new XlsxContainer();
        xlsxContainer.setNameFile(fileExcel.getName());
        OPCPackage file = OPCPackage.open(fileExcel);//открываем документ
        xlsxContainer.setWorkbook(new XSSFWorkbook(file));
        xlsxContainer.setSheet(xlsxContainer.getWorkbook().getSheetAt(index-1));
        return xlsxContainer;
    }
}
