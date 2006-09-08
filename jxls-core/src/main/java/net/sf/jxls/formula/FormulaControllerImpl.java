package net.sf.jxls.formula;

import org.apache.poi.hssf.util.CellReference;
import net.sf.jxls.tag.Point;
import net.sf.jxls.transformation.BlockTransformation;
import net.sf.jxls.transformation.DuplicateTransformation;
import net.sf.jxls.transformer.Sheet;
import net.sf.jxls.transformer.Workbook;
import net.sf.jxls.util.SheetHelper;
import net.sf.jxls.util.Util;

import java.util.Iterator;
import java.util.List;
import java.util.Set;

/**
 * @author Leonid Vysochyn
 */
public class FormulaControllerImpl implements FormulaController {
    protected  static String leftReplacementMarker = "{";
    protected  static String rightReplacementMarker = "}";
    protected  static String regexReplacementMarker = "\\" + leftReplacementMarker + "[(),a-zA-Z0-9_ :*+/.-]+" + "\\" + rightReplacementMarker;

    /**
     * Ref cell in a formula string is replaced with result cell enclosed with replacement markers to be able not to replace
     * already replaced cells
     * @param formulaPart
     * @param refCell
     * @param newCell
     * @return updated formula string
     */
    public static String replaceFormulaPart(String formulaPart, String refCell, String newCell) {
        String replacedFormulaPart = "";
        String[] parts = formulaPart.split(regexReplacementMarker, 2);
        for(; parts.length == 2; parts = formulaPart.split(regexReplacementMarker, 2) ){
            replacedFormulaPart += parts[0].replaceAll( refCell, leftReplacementMarker + newCell + rightReplacementMarker );
            int secondPartIndex;
            if( parts[1].length() != 0 ){
                secondPartIndex = formulaPart.indexOf(parts[1], parts[0].length());
            }else{
                secondPartIndex = formulaPart.length();
            }
            replacedFormulaPart += formulaPart.substring( parts[0].length(), secondPartIndex );
            formulaPart = parts[1];
        }
        replacedFormulaPart += parts[0].replaceAll( refCell, leftReplacementMarker + newCell + rightReplacementMarker );
        return replacedFormulaPart;
    }


    Workbook workbook;

    public FormulaControllerImpl(Workbook workbook) {
        this.workbook = workbook;
    }

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
                            updatedFormula = replaceFormulaPart(updatedFormula, cellRef, newCell);
                            isFormulaUpdated = true;
                        }else if( resultCells.size() > 1 ){
                            String refSheetName = extractRefSheetName( cellRef );
                            String newCell = detectCellRange( refSheetName, resultCells );
                            updatedFormula = replaceFormulaPart(formula.getFormula(), cellRef, newCell);
                            isFormulaUpdated = true;
                        }
                    }
                }
                // remove replacement markers
                updatedFormula = updatedFormula.replaceAll( "\\" + leftReplacementMarker, "" );
                updatedFormula = updatedFormula.replaceAll( "\\" + rightReplacementMarker, "" );
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

    protected static final String regexCellCharPart = "[0-9]+";
    protected static final String regexCellDigitPart = "[a-zA-Z]+";
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
            String firstCellCharPart = firstCell.split(regexCellCharPart)[0];
            String firstCellDigitPart = firstCell.split(regexCellDigitPart)[1];
            int cellNumber = Integer.parseInt( firstCellDigitPart );
            String nextCell, cellCharPart, cellDigitPart;
            for (int i = 1; i < cells.size() && isColumnRange; i++) {
                nextCell = (String) cells.get(i);
                cellCharPart = nextCell.split( regexCellCharPart )[0];
                cellDigitPart = nextCell.split( regexCellDigitPart )[1];
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
            String firstCellDigitPart = firstCell.split(regexCellDigitPart)[1];
            String nextCell, cellDigitPart;
            CellReference cellRef = new CellReference( firstCell );
            int cellNumber = cellRef.getCol();
            for (int i = 1; i < cells.size() && isRowRange; i++) {
                nextCell = (String) cells.get(i);
                cellDigitPart = nextCell.split( regexCellDigitPart )[1];
                cellRef = new CellReference( nextCell );
                if( !firstCellDigitPart.equalsIgnoreCase( cellDigitPart ) || cellRef.getCol() != ++cellNumber ){
                    isRowRange = false;
                }
            }
        }
        return isRowRange;
    }


}
