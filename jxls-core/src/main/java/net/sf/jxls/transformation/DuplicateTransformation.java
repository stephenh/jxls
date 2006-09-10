package net.sf.jxls.transformation;

import net.sf.jxls.tag.Block;
import net.sf.jxls.tag.Point;
import org.apache.poi.hssf.util.CellReference;

import java.util.List;
import java.util.ArrayList;

/**
 * Defines duplicate transformation for {@link Block}
 * @author Leonid Vysochyn
 */
public class DuplicateTransformation extends BlockTransformation {

    int duplicateNumber;

    public DuplicateTransformation(Block block, int duplicateNumber) {
        super(block);
        this.duplicateNumber = duplicateNumber;
    }

    public Block getBlockAfterTransformation() {
        return null;
    }

    public List transformCell(Point p) {
        List cells;
        if( block.contains( p ) ){
            cells = new ArrayList();
            Point rp = p;
            cells.add( p );
            for( int i = 0; i < duplicateNumber; i++){
                cells.add( rp = rp.shift( block.getNumberOfRows(), 0));
            }
        }else{
            cells = new ArrayList();
            cells.add( p );
        }
        return cells;
    }

    public String getDuplicatedCellRef(String sheetName, String cell, int duplicateBlock){
        CellReference cellRef = new CellReference(cell);
        int rowNum = cellRef.getRow();
        short colNum = cellRef.getCol();
        String refSheetName = cellRef.getSheetName();
        String resultCellRef = cell;
        if( block.getSheet().getSheetName().equalsIgnoreCase( refSheetName ) || (refSheetName == null && block.getSheet().getSheetName().equalsIgnoreCase( sheetName ))){
            // sheet check passed
            Point p = new Point( rowNum, colNum );
            if( block.contains( p ) && duplicateNumber >= 1 && duplicateNumber >= duplicateBlock){
                p = p.shift( block.getNumberOfRows() * duplicateBlock, 0 );
                resultCellRef = p.toString( refSheetName );
            }
        }
        return resultCellRef;
    }

    public List transformCell(String sheetName, String cell) {
        CellReference cellRef = new CellReference(cell);
        int rowNum = cellRef.getRow();
        short colNum = cellRef.getCol();
        String refSheetName = cellRef.getSheetName();
        List cells = new ArrayList();
        if( block.getSheet().getSheetName().equalsIgnoreCase( refSheetName ) || (refSheetName == null && block.getSheet().getSheetName().equalsIgnoreCase( sheetName ))){
            // sheet check passed
            Point p = new Point( rowNum, colNum );
            if( block.contains( p ) /*&& duplicateNumber >= 1*/){
                cells.add( p.toString(refSheetName) );
                for( int i = 0; i < duplicateNumber; i++){
                    p = p.shift( block.getNumberOfRows(), 0);
                    cells.add( p.toString( refSheetName ));
                }
            }
        }
        return cells;
    }

    public boolean equals(Object obj) {
        if( obj != null && obj instanceof DuplicateTransformation ){
            DuplicateTransformation dt = (DuplicateTransformation) obj;
            return ( super.equals( obj ) && dt.duplicateNumber == duplicateNumber);
        }else{
            return false;
        }
    }

    public int hashCode() {
        int result = super.hashCode();
        result = 29 * result + duplicateNumber;
        return result;
    }

    public String toString() {
        return "DuplicateTransformation: {" + super.toString() + ", duplicateNumber=" + duplicateNumber + "}";
    }
}
