package net.sf.jxls.transformer;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import net.sf.jxls.transformer.Sheet;
import net.sf.jxls.util.SheetHelper;

import java.util.List;
import java.util.ArrayList;
import java.util.Map;
import java.util.HashMap;

/**
 * Represents excel workbook
 * @author Leonid Vysochyn
 */
public class Workbook {
    List sheets = new ArrayList();
    /**
     * POI Excel workbook object
     */
    HSSFWorkbook hssfWorkbook;


    Configuration configuration = new Configuration();

    public Workbook(HSSFWorkbook hssfWorkbook) {
        this.hssfWorkbook = hssfWorkbook;
    }

    public Workbook(HSSFWorkbook hssfWorkbook, Configuration configuration) {
        this.hssfWorkbook = hssfWorkbook;
        this.configuration = configuration;
    }

    public Workbook(HSSFWorkbook hssfWorkbook, List sheets) {
        this.hssfWorkbook = hssfWorkbook;
        this.sheets = sheets;
    }

    public Workbook(HSSFWorkbook hssfWorkbook, List sheets, Configuration configuration) {
        this.hssfWorkbook = hssfWorkbook;
        this.sheets = sheets;
        this.configuration = configuration;
    }

    public HSSFWorkbook getHssfWorkbook() {
        return hssfWorkbook;
    }

    public void setHssfWorkbook(HSSFWorkbook hssfWorkbook) {
        this.hssfWorkbook = hssfWorkbook;
    }

    public void addSheet(Sheet sheet){
        sheets.add( sheet );
        sheet.setWorkbook( this );
    }

    public Map getListRanges(){
        Map listRanges = new HashMap();
        for (int i = 0; i < sheets.size(); i++) {
            Sheet sheet = (Sheet) sheets.get(i);
            listRanges.putAll( sheet.getListRanges() );
        }
        return listRanges;
    }

    public List getFormulas(){
        List formulas = new ArrayList();
        for (int i = 0; i < sheets.size(); i++) {
            Sheet sheet = (Sheet) sheets.get(i);
            formulas.addAll( SheetHelper.findFormulas( sheet ) );
        }
        return formulas;
    }

    public List getSheets() {
        return sheets;
    }

    public int getNumberOfSheets(){
        return sheets.size();
    }

    public Sheet getSheetAt(int sheetNo){
        return (Sheet) sheets.get( sheetNo );
    }

}
