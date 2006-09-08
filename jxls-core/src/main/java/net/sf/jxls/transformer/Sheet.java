package net.sf.jxls.transformer;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFRow;
import net.sf.jxls.formula.Formula;
import net.sf.jxls.formula.ListRange;
import net.sf.jxls.parser.Cell;
import net.sf.jxls.util.SheetHelper;
import net.sf.jxls.tag.Block;
import net.sf.jxls.tag.Point;

import java.util.*;

/**
 * Represents excel worksheet 
 * @author Leonid Vysochyn
 */
public class Sheet {

    Workbook workbook;

    /**
     * POI Excel workbook object
     */
    HSSFWorkbook hssfWorkbook;

    /**
     * POI Excel sheet representation
     */
    HSSFSheet hssfSheet;
    /**
     * This variable stores all list ranges found while processing template file
     */
    private Map listRanges = new HashMap();
    /**
     * Stores all named HSSFCell objects
     */
    private Map namedCells = new HashMap();

    Configuration configuration = new Configuration();

    public Sheet() {
    }

    public Sheet(HSSFWorkbook hssfWorkbook, HSSFSheet hssfSheet, Configuration configuration) {
        this.hssfWorkbook = hssfWorkbook;
        this.hssfSheet = hssfSheet;
        this.configuration = configuration;
    }

    public Sheet(HSSFWorkbook hssfWorkbook, HSSFSheet hssfSheet) {
        this.hssfWorkbook = hssfWorkbook;
        this.hssfSheet = hssfSheet;
    }

    public String getSheetName(){
        for(int i = 0; i < hssfWorkbook.getNumberOfSheets(); i++){
            HSSFSheet sheet = hssfWorkbook.getSheetAt( i );
            if( sheet == hssfSheet ){
                return hssfWorkbook.getSheetName( i );
            }
        }
        return null;
    }

    public HSSFWorkbook getHssfWorkbook() {
        return hssfWorkbook;
    }

    public void setHssfWorkbook(HSSFWorkbook hssfWorkbook) {
        this.hssfWorkbook = hssfWorkbook;
    }

    public void setHssfSheet(HSSFSheet hssfSheet) {
        this.hssfSheet = hssfSheet;
    }

    public HSSFSheet getHssfSheet() {
        return hssfSheet;
    }

    public Configuration getConfiguration() {
        return configuration;
    }

    public Map getListRanges() {
        return listRanges;
    }

    public Map getNamedCells() {
        return namedCells;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public void addNamedCell(String label, Cell cell){
        namedCells.put( label, cell );
    }

    public void addListRange(String name, ListRange range){
        listRanges.put( name, range );
    }

    public int getMaxColNum(){
        int maxColNum = 0;
        for(int i = hssfSheet.getFirstRowNum(); i <= hssfSheet.getLastRowNum(); i++){
            HSSFRow hssfRow = hssfSheet.getRow( i );
            if( hssfRow != null ){
                if( hssfRow.getLastCellNum() > maxColNum ){
                    maxColNum = hssfRow.getLastCellNum();
                }
            }
        }
        return maxColNum;
    }

    public Map getFormulaCellRefsToUpdate(Block block) {
        List formulas = SheetHelper.findFormulas( this );
        Map formulaCellRefsToUpdate = new HashMap();
        String transformedSheetName = block.getSheet().getSheetName();
        for (int j = 0; j < formulas.size(); j++) {
            Formula formula = (Formula) formulas.get(j);
            Set refCells = formula.findRefCells();
            Point key = new Point( formula.getRowNum().intValue(), formula.getCellNum().shortValue() );
            List refCellsToUpdate = new ArrayList();
            for (Iterator iterator = refCells.iterator(); iterator.hasNext();) {
                String refCell = (String) iterator.next();
                if( refCell.indexOf("!")<0 ){
                    if( block.getSheet().getHssfSheet() == hssfSheet ){
                        updateRefCellList(refCellsToUpdate, refCell, block);
                    }
                }else{
                    int index = refCell.indexOf("!");
                    String sheetName = refCell.substring(0, index);
                    if( sheetName.equalsIgnoreCase( transformedSheetName ) ){
                        updateRefCellList(refCellsToUpdate, refCell, block);
                    }
                }
            }
            if( !refCellsToUpdate.isEmpty() ){
                formulaCellRefsToUpdate.put( key, refCellsToUpdate );
            }
        }
        return formulaCellRefsToUpdate;
    }

    private void updateRefCellList(List refCellsToUpdate, String refCell, Block block) {
        Point point = new Point(refCell);
        if( block.contains( point ) ){
            refCellsToUpdate.add( refCell );
        }
    }


}
