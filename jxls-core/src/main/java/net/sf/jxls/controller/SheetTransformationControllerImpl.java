package net.sf.jxls.controller;

import org.apache.poi.hssf.usermodel.HSSFRow;
import net.sf.jxls.util.Util;
import net.sf.jxls.util.SheetHelper;
import net.sf.jxls.util.TagBodyHelper;
import net.sf.jxls.formula.*;
import net.sf.jxls.transformation.DuplicateTransformation;
import net.sf.jxls.transformation.RemoveTransformation;
import net.sf.jxls.transformation.ShiftTransformation;
import net.sf.jxls.controller.SheetTransformationController;
import net.sf.jxls.tag.Block;
import net.sf.jxls.tag.Point;
import net.sf.jxls.transformer.RowCollection;
import net.sf.jxls.transformer.Sheet;
import net.sf.jxls.parser.Cell;

import java.util.*;

/**
 * This class controls and saves all transforming operations on spreadsheet cells.
 * It implements {@link net.sf.jxls.controller.SheetTransformationController} interface using special {@link TransformationMatrix}
 * to track all cells transformations
 * @author Leonid Vysochyn
 */
public class SheetTransformationControllerImpl implements SheetTransformationController {

    List transformations = new ArrayList();

    Sheet sheet;
    TagBodyHelper helper;
    FormulaController formulaController;

    public SheetTransformationControllerImpl(Sheet sheet) {
        this.sheet = sheet;
        helper = new TagBodyHelper();
        formulaController = new FormulaControllerImpl( sheet.getWorkbook() );
    }

    public int duplicateDown( Block block, int n ){
        if( n > 0 ){
            ShiftTransformation shiftTransformation = new ShiftTransformation(new Block(sheet, block.getEndRowNum() + 1, Integer.MAX_VALUE), n * block.getNumberOfRows(), 0);
            transformations.add( shiftTransformation);
            if( block.getSheet() == null ){
                block.setSheet( sheet );
            }
            DuplicateTransformation duplicateTransformation = new DuplicateTransformation(block, n);
            transformations.add( duplicateTransformation );
            Map formulaRefCellUpdates = findFormulaRefCellsToUpdate(block);
            formulaController.updateFormulas( shiftTransformation );
            formulaController.updateFormulas( duplicateTransformation );
            return TagBodyHelper.duplicateDown( sheet.getHssfSheet(), block, n, formulaRefCellUpdates );
        }else{
            return 0;
        }
    }

    private Map findFormulaRefCellsToUpdate(Block block) {
        Map formulaCellRefUpdates = new HashMap();
        List formulas = SheetHelper.findFormulas( sheet, block );
        for (int i = 0; i < formulas.size(); i++) {
            Formula formula = (Formula) formulas.get(i);
            Set refCells = formula.findRefCells();

            Point key = new Point( formula.getRowNum().intValue(), formula.getCellNum().shortValue() );
            List refCellsToUpdate = new ArrayList();
            for (Iterator iterator = refCells.iterator(); iterator.hasNext();) {
                String refCell = (String) iterator.next();
                if( refCell.indexOf("!")<0 ){
                    Point point = new Point(refCell);
                    if( block.contains( point ) ){
                        refCellsToUpdate.add( refCell );
                    }
                    if( !refCellsToUpdate.isEmpty() ){
                        formulaCellRefUpdates.put( key, refCellsToUpdate );
                    }
                }
            }
        }
        return formulaCellRefUpdates;
    }

    public int duplicateRight(Block block, int n) {
        if( n > 0 ){
            return TagBodyHelper.duplicateRight( sheet.getHssfSheet(), block, n) ;
        }else{
            return 0;
        }
    }

    public void removeBorders(Block block) {
        transformations.add( new RemoveTransformation( new Block(sheet, block.getStartRowNum(), block.getStartRowNum())));
        ShiftTransformation shiftTransformation1 = new ShiftTransformation(new Block(sheet, block.getStartRowNum() + 1, Integer.MAX_VALUE), -1, 0);
        transformations.add( shiftTransformation1 );
        transformations.add( new RemoveTransformation( new Block(sheet, block.getEndRowNum() - 1, block.getEndRowNum() - 1 ) ));
        ShiftTransformation shiftTransformation2 = new ShiftTransformation(new Block(sheet, block.getEndRowNum(), Integer.MAX_VALUE), -1, 0);
        transformations.add( shiftTransformation2 );
        formulaController.updateFormulas( shiftTransformation1 );
        formulaController.updateFormulas( shiftTransformation2 );
        TagBodyHelper.removeBorders( sheet.getHssfSheet(), block );

    }

    public void removeLeftRightBorders(Block block) {
        TagBodyHelper.removeLeftRightBorders( sheet.getHssfSheet(), block);
    }

    public void removeRowCells(HSSFRow row, short startCellNum, short endCellNum) {
        TagBodyHelper.removeRowCells( sheet.getHssfSheet(), row, startCellNum, endCellNum );
    }

    public void removeBodyRows(Block block) {
        transformations.add( new RemoveTransformation( block ) );
        ShiftTransformation shiftTransformation = new ShiftTransformation(new Block(sheet, block.getEndRowNum() + 1, Integer.MAX_VALUE), -block.getNumberOfRows(), 0);
        transformations.add( shiftTransformation );
        formulaController.updateFormulas( shiftTransformation );
        TagBodyHelper.removeBodyRows( sheet.getHssfSheet(), block );
    }


    public void duplicateRow(RowCollection rowCollection) {
        int startRowNum = rowCollection.getParentRow().getHssfRow().getRowNum();
        int endRowNum = startRowNum + rowCollection.getDependentRowNumber();
        ShiftTransformation shiftTransformation = new ShiftTransformation(new Block(sheet, endRowNum + 1, Integer.MAX_VALUE), rowCollection.getCollectionProperty().getCollection().size() - 1, 0);
        transformations.add( shiftTransformation);
        DuplicateTransformation duplicateTransformation = new DuplicateTransformation(new Block(sheet, startRowNum, endRowNum), rowCollection.getCollectionProperty().getCollection().size()-1);
        transformations.add( duplicateTransformation );
        List cells = rowCollection.getRowCollectionCells();
        for (int i = 0; i < cells.size(); i++) {
            Cell cell = (Cell) cells.get(i);
            if( cell!= null && cell.getHssfCell() != null){
            }
        }
        formulaController.updateFormulas( shiftTransformation );
        formulaController.updateFormulas( duplicateTransformation );
        Util.duplicateRow( rowCollection );
    }

    public List getTransformations() {
        return transformations;
    }

    public Sheet getSheet() {
        return sheet;
    }
}
