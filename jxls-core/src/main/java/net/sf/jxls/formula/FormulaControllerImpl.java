package net.sf.jxls.formula;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import net.sf.jxls.tag.Point;
import net.sf.jxls.transformation.BlockTransformation;
import net.sf.jxls.transformation.DuplicateTransformation;
import net.sf.jxls.transformer.Sheet;
import net.sf.jxls.transformer.Workbook;
import net.sf.jxls.util.SheetHelper;
import net.sf.jxls.util.Util;

import java.util.*;

/**
 * @author Leonid Vysochyn
 */
public class FormulaControllerImpl implements FormulaController {
    protected final Log log = LogFactory.getLog(getClass());

    protected Map sheetFormulasMap;


    Workbook workbook;

    public FormulaControllerImpl(Workbook workbook) {
        this.workbook = workbook;
        sheetFormulasMap = workbook.createFormulaSheetMap();
    }

    public void updateWorkbookFormulas(BlockTransformation transformation){
        Set sheetNames = sheetFormulasMap.keySet();
        for (Iterator iterator = sheetNames.iterator(); iterator.hasNext();) {
            String sheetName =  (String) iterator.next();
            List formulas = (List) sheetFormulasMap.get( sheetName );
            for (int i = 0; i < formulas.size(); i++) {
                Formula formula = (Formula) formulas.get(i);
                Set cellRefs = formula.getCellRefs();
                for (Iterator iter = cellRefs.iterator(); iter.hasNext();) {
                    CellRef cellRef = (CellRef) iter.next();
                    if( !(transformation instanceof DuplicateTransformation && transformation.getBlock().contains(new Point( cellRef.toString() )) &&
                            transformation.getBlock().contains( formula.getRowNum().intValue(), formula.getCellNum().intValue())) ){
                        List resultCells = transformation.transformCell( sheetName, cellRef.toString() );
                        if( resultCells.size() == 1 ){
                            String newCell = (String) resultCells.get(0);
                            cellRef.update( newCell );
                        }else if( resultCells.size() > 1 ){
                            cellRef.update( resultCells );
                        }
                    }
                }
                formula.updateReplacedRefCellsCollection();
                if( formula.getSheet().getSheetName().equals( transformation.getBlock().getSheet().getSheetName() ) ){
                    Point p = new Point( formula.getRowNum().intValue(), formula.getCellNum().shortValue() );
                    List points = transformation.transformCell( p );
                    if( points!=null && !points.isEmpty()){
                        if(points.size() == 1){
                            Point newPoint = (Point) points.get(0);
                            formula.setRowNum( new Integer( newPoint.getRow() ));
                            formula.setCellNum( new Integer( newPoint.getCol() ));
                        }else{
                            List sheetFormulas = (List) sheetFormulasMap.get( formula.getSheet().getSheetName() );
                            for (int j = 1; j < points.size(); j++) {
                                Point point = (Point) points.get(j);
                                Formula newFormula = new Formula( formula );
                                newFormula.setRowNum( new Integer(point.getRow()) );
                                newFormula.setCellNum( new Integer(point.getCol() ) );
                                Set newCellRefs = newFormula.getCellRefs();
                                for (Iterator iterator1 = newCellRefs.iterator(); iterator1.hasNext();) {
                                    CellRef newCellRef =  (CellRef) iterator1.next();
                                    if( transformation.getBlock().contains( new Point( newCellRef.toString() ) ) && transformation.getBlock().contains( p ) ){
                                        newCellRef.update(((DuplicateTransformation)transformation).getDuplicatedCellRef( sheetName, newCellRef.toString(), j));
                                    }
                                }
                                sheetFormulas.add( newFormula );
                            }
                        }
                    }
                }
            }
        }
    }

    public Map getSheetFormulasMap() {
        return sheetFormulasMap;
    }

    public void writeFormulas(FormulaResolver formulaResolver) {
        Set sheetNames = sheetFormulasMap.keySet();
        for (Iterator iterator = sheetNames.iterator(); iterator.hasNext();) {
            String sheetName =  (String) iterator.next();
            List formulas = (List) sheetFormulasMap.get( sheetName );
            for (int i = 0; i < formulas.size(); i++) {
                Formula formula = (Formula) formulas.get(i);
                String formulaString = formulaResolver.resolve( formula, null);
                HSSFRow hssfRow = formula.getSheet().getHssfSheet().getRow(formula.getRowNum().intValue());
                HSSFCell hssfCell = hssfRow.getCell(formula.getCellNum().shortValue());
                if (formulaString != null) {
                    if( hssfCell == null ){
                        hssfCell = hssfRow.createCell( formula.getCellNum().shortValue() );
                    }
                    try {
                        hssfCell.setCellFormula(formulaString);
                    } catch (RuntimeException e) {
                        log.error("Can't set formula: " + formulaString, e);
                        throw new RuntimeException("Can't set formula: " + formulaString, e );
                    }
                }
            }
        }
    }

    // todo: to deprecate
    public void updateFormulas(BlockTransformation transformation) {
        List sheets = workbook.getSheets();
        for (int i = 0; i < sheets.size(); i++) {
            Sheet sheet = (Sheet) sheets.get(i);
            List formulas = SheetHelper.findFormulas( sheet );
            String sheetName = sheet.getSheetName();
            for (int j = 0; j < formulas.size(); j++) {
                Formula formula = (Formula) formulas.get(j);
                Set refCells = formula.findRefCells();
                String updatedFormula = formula.getFormula();
                boolean isFormulaUpdated = false;
                for (Iterator iterator = refCells.iterator(); iterator.hasNext();) {
                    String cellRef = (String) iterator.next();
                    if( !(transformation instanceof DuplicateTransformation && transformation.getBlock().contains(new Point( cellRef )) &&
                            transformation.getBlock().contains( formula.getRowNum().intValue(), formula.getCellNum().intValue())) ){
                        List resultCells = transformation.transformCell( sheetName, cellRef );
                        if( resultCells.size() == 1 ){
                            String newCell = (String) resultCells.get(0);
                            updatedFormula = CellRef.replaceFormulaPart(updatedFormula, cellRef, newCell);
                            isFormulaUpdated = true;
                        }else if( resultCells.size() > 1 ){
                            String refSheetName = extractRefSheetName( cellRef );
                            String newCell = detectCellRange( refSheetName, resultCells );
                            updatedFormula = CellRef.replaceFormulaPart(formula.getFormula(), cellRef, newCell);
                            isFormulaUpdated = true;
                        }
                    }
                }
                // remove replacement markers
                updatedFormula = updatedFormula.replaceAll( "\\" + CellRef.leftReplacementMarker, "" );
                updatedFormula = updatedFormula.replaceAll( "\\" + CellRef.rightReplacementMarker, "" );
//                formula.setFormula( updatedFormula );
                if( isFormulaUpdated ){
                    Util.updateCellValue( sheet.getHssfSheet(), formula.getRowNum().intValue(),
                            formula.getCellNum().shortValue(), sheet.getConfiguration().getStartFormulaToken() + updatedFormula
                    + sheet.getConfiguration().getEndFormulaToken());
                }
            }
        }
    }

    String getSheetName(String cell){
        CellReference cellRef = new CellReference( cell );
        return cellRef.getSheetName();
    }

    protected String cellRangeSeparator = ":";

    private String extractRefSheetName(String refCell) {
        if( refCell != null ){
            if( refCell.indexOf("!") < 0 ){
                return null;
            }else{
                return refCell.substring(0, refCell.indexOf("!") );
            }
        }
        return null;
    }

    private String extractCellName(String refCell) {
        if( refCell != null ){
            if( refCell.indexOf("!") < 0 ){
                return refCell;
            }else{
                return refCell.substring( refCell.indexOf("!") + 1 );
            }
        }
        return null;
    }


    String detectCellRange(String refSheetName, List cells) {
        cutSheetRefFromCells( cells );
        String firstCell = (String) cells.get( 0 );
        String range = firstCell;
        if( firstCell != null && firstCell.length() > 0 ){
            if( isRowRange(cells) || isColumnRange(cells) ){
                String lastCell = (String) cells.get( cells.size() - 1 );
                range = getRefCellName(refSheetName, firstCell) + cellRangeSeparator + lastCell.toUpperCase();
            }else{
                range = buildCommaSeparatedListOfCells(refSheetName, cells );
            }
        }
        return range;
    }

    private void cutSheetRefFromCells(List cells) {
        for (int i = 0; i < cells.size(); i++) {
            String cell = (String) cells.get(i);
            cells.set(i, extractCellName( cell ) );
        }
    }

    String buildCommaSeparatedListOfCells(String refSheetName, List cells) {
        String listOfCells = "";
        for (int i = 0; i < cells.size() - 1; i++) {
            String cell = (String) cells.get(i);
            listOfCells += getRefCellName(refSheetName, cell) + ",";
        }
        listOfCells += getRefCellName( refSheetName, (String) cells.get( cells.size() - 1 ));
        return listOfCells;
    }


    String getRefCellName(String refSheetName, String cellName){
        if( refSheetName == null ){
            return cellName.toUpperCase();
        }else{
            return refSheetName + "!" + cellName.toUpperCase();
        }
    }

    boolean isColumnRange(List cells) {
        String firstCell = (String) cells.get( 0 );
        boolean isColumnRange = true;
        if( firstCell != null && firstCell.length() > 0 ){
            String firstCellCharPart = firstCell.split(CellRef.regexCellCharPart)[0];
            String firstCellDigitPart = firstCell.split(CellRef.regexCellDigitPart)[1];
            int cellNumber = Integer.parseInt( firstCellDigitPart );
            String nextCell, cellCharPart, cellDigitPart;
            for (int i = 1; i < cells.size() && isColumnRange; i++) {
                nextCell = (String) cells.get(i);
                cellCharPart = nextCell.split( CellRef.regexCellCharPart )[0];
                cellDigitPart = nextCell.split( CellRef.regexCellDigitPart )[1];
                if( !firstCellCharPart.equalsIgnoreCase( cellCharPart ) || Integer.parseInt(cellDigitPart) != ++cellNumber ){
                    isColumnRange = false;
                }
            }
        }
        return isColumnRange;
    }

    boolean isRowRange(List cells) {
        String firstCell = (String) cells.get( 0 );
        boolean isRowRange = true;
        if( firstCell != null && firstCell.length() > 0 ){
            String firstCellDigitPart = firstCell.split(CellRef.regexCellDigitPart)[1];
            String nextCell, cellDigitPart;
            CellReference cellRef = new CellReference( firstCell );
            int cellNumber = cellRef.getCol();
            for (int i = 1; i < cells.size() && isRowRange; i++) {
                nextCell = (String) cells.get(i);
                cellDigitPart = nextCell.split( CellRef.regexCellDigitPart )[1];
                cellRef = new CellReference( nextCell );
                if( !firstCellDigitPart.equalsIgnoreCase( cellDigitPart ) || cellRef.getCol() != ++cellNumber ){
                    isRowRange = false;
                }
            }
        }
        return isRowRange;
    }


}
