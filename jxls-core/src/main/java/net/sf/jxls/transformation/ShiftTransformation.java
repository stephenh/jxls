package net.sf.jxls.transformation;

import net.sf.jxls.transformation.BlockTransformation;
import net.sf.jxls.tag.Block;
import net.sf.jxls.tag.Point;
import org.apache.poi.hssf.util.CellReference;

import java.util.List;
import java.util.ArrayList;

/**
 * Defines simple shift transformation
 * @author Leonid Vysochyn
 */
public class ShiftTransformation extends BlockTransformation {
    int rowShift;
    int colShift;

    public ShiftTransformation(Block block, int rowShift, int colShift) {
        super(block);
        this.rowShift = rowShift;
        this.colShift = colShift;
    }

    public Block getBlockAfterTransformation() {
        return null;  //To change body of implemented methods use File | Settings | File Templates.
    }

    public List transformCell(Point p) {
        List cells = new ArrayList();
        if( block.contains( p ) || block.isAbove( p )){
            cells.add( p.shift( rowShift, colShift ) );
        }else{
            cells.add( p );
        }
        return cells;
    }

    public List transformCell(String sheetName, String cell) {
        CellReference cellRef = new CellReference(cell);
        int rowNum = cellRef.getRow();
        short colNum = cellRef.getCol();
        String refSheetName = cellRef.getSheetName();
        List cells = new ArrayList();
        if( block.getSheet().getSheetName().equalsIgnoreCase( refSheetName ) || (refSheetName == null && block.getSheet().getSheetName().equalsIgnoreCase( sheetName ))){
            Point p = new Point( rowNum, colNum );
            if( block.contains( p ) || block.isAbove( p ) ){
                p = p.shift( rowShift, colShift );
                cellRef = new CellReference( p.getRow(), p.getCol() );
                if( refSheetName != null ){
                    cells.add( refSheetName + "!" + cellRef.toString());
                }else{
                    cells.add( cellRef.toString() );
                }
            }
        }
        return cells;
    }

    public boolean equals(Object obj) {
        if( obj != null && obj instanceof ShiftTransformation ){
            ShiftTransformation st = (ShiftTransformation) obj;
            return ( super.equals( obj ) && rowShift == st.rowShift && colShift == st.colShift);
        }else{
            return false;
        }
    }

    public int hashCode() {
        int result = super.hashCode();
        result = 29 * result + rowShift;
        result = 29 * result + colShift;
        return result;
    }

    public String toString() {
        return "ShiftTransformation: {" + super.toString() + ", shift=(" + rowShift + ", " + colShift + ")}";
    }
}
