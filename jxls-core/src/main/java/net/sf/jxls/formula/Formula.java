package net.sf.jxls.formula;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import net.sf.jxls.transformer.Sheet;
import net.sf.jxls.parser.Cell;
import net.sf.jxls.controller.SheetCellFinder;

import java.util.*;
import java.util.regex.Pattern;
import java.util.regex.Matcher;


/**
 * Represents formula cell
 * @author Leonid Vysochyn
 */
public class Formula {
    protected final Log log = LogFactory.getLog(getClass());

    private String formula;
    private Integer rowNum;
    private Integer cellNum;
    static final String inlineFormulaToken = "#";
    static final String formulaListRangeToken = "@";
    private String adjustedFormula;

    private Sheet sheet;

    public Sheet getSheet() {
        return sheet;
    }

    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    public Formula(String formula) {
        this.formula = formula;
    }

    public Formula() {
    }

    public String getFormula() {
        return formula;
    }

    public void setFormula(String formula) {
        this.formula = formula;
    }

    public Integer getRowNum() {
        return rowNum;
    }

    public void setRowNum(Integer rowNum) {
        this.rowNum = rowNum;
    }

    public Integer getCellNum() {
        return cellNum;
    }

    public void setCellNum(Integer cellNum) {
        this.cellNum = cellNum;
    }

    public boolean isInline() {
        return formula.indexOf(inlineFormulaToken) >= 0;
    }

    public String getInlineFormula(int n) {
        if (isInline()) {
            return formula.replaceAll(inlineFormulaToken, Integer.toString(n));
        } else {
            return formula;
        }
    }

    /**
     * @return Formula string that should be set into Excel cell using POI
     */
    public String getAppliedFormula(Map listRanges, Map namedCells) {
        String codedFormula = formula;
        String appliedFormula = "";
        String delimiter = formulaListRangeToken;
        int index = codedFormula.indexOf(delimiter);
        boolean isExpression = false;
        while (index >= 0) {
            String token = codedFormula.substring(0, index);
            if (isExpression) {
                // this is formula coded expression variable
                // look into the listRanges to see do we have cell range for it
                if (listRanges.containsKey(token)) {
                    appliedFormula += ((ListRange) listRanges.get(token)).toExcelCellRange();
                } else if (namedCells.containsKey(token)) {
                    appliedFormula += ((Cell) namedCells.get(token)).toCellName();
                } else {
                    log.warn("can't find list range or named cell for " + token);
                    // returning null if we don't have given list range or named cell so we don't need to set formula to avoid error
                    return null;
                }
            } else {
                appliedFormula += token;
            }
            codedFormula = codedFormula.substring(index + 1);
            index = codedFormula.indexOf(delimiter);
            isExpression = !isExpression;
        }
        appliedFormula += codedFormula;
        return appliedFormula;
    }

    String adjust(SheetCellFinder cellFinder){
//        String adjustedFormula = formula;
//        Set refCells = findRefCells();
//        for (Iterator iterator = refCells.iterator(); iterator.hasNext();) {
//            String refCell = (String) iterator.next();
//            String newCell = cellFinder.findCell( refCell );
//            adjustedFormula = adjustedFormula.replaceAll( refCell, newCell );
//        }
//        formula = adjustedFormula;
        return formula;
    }

    private static final String regexCellRef = "([a-zA-Z]+[a-zA-Z0-9]*![a-zA-Z]+[0-9]+|[a-zA-Z]+[0-9]+|'[^?\\\\/:'*]+'![a-zA-Z]+[0-9]+)";
    private static final Pattern regexCellRefPattern = Pattern.compile( regexCellRef );

    public Set findRefCells() {
        Set refCells = new HashSet();
        Matcher refCellMatcher = regexCellRefPattern.matcher( formula );
        while( refCellMatcher.find() ){
            refCells.add( refCellMatcher.group() );
        }
        return refCells;
    }


    public String toString() {
        return "Formula{" +
                "formula='" + formula + "'" +
                ", rowNum=" + rowNum +
                ", cellNum=" + cellNum +
                "}";
    }

    public boolean containsListRanges() {
        return formula.indexOf( formulaListRangeToken ) >= 0;
    }

}
