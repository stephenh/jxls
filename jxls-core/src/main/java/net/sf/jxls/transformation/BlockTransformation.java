package net.sf.jxls.transformation;

import net.sf.jxls.tag.Block;
import net.sf.jxls.tag.Point;

import java.util.List;

/**
 * This class defines common {@link net.sf.jxls.tag.Block} transformation behaviour
 * @author Leonid Vysochyn
 */
public abstract class BlockTransformation {

    protected Block block;

    /**
     * @param block - defines transformation {@link net.sf.jxls.tag.Block}
     */
    protected BlockTransformation(Block block) {
        this.block = block;
    }


    public Block getBlock() {
        return block;
    }

    public void setBlock(Block block) {
        this.block = block;
    }

    boolean contains(int row, int col){
        return block.contains( row, col);
    }

    boolean contains(Point p){
        return block.contains( p );
    }


    public abstract Block getBlockAfterTransformation();

    /**
     * Transforms given spreadsheet cell
     * @param p - {@link net.sf.jxls.tag.Point} representing spreadsheet cell to transform
     * @return {@link List} of {@link Point} objects which are result of source cell transformation
     */
    public abstract List transformCell(Point p);

    /**
     * Transforms given spreadsheet cell
     * @param cell - cell name reference representing spreadsheet cell to transform
     * @return {@link List} of cell names which are result of source cell transformation
     */
    public abstract List transformCell(String sheetName, String cell);

    public boolean equals(Object obj) {
        if( obj != null && obj instanceof BlockTransformation ){
            BlockTransformation bt = (BlockTransformation) obj;
            return ((block!=null && block.equals( bt.block )) || (block == null && bt.block == null));
        }else{
            return false;
        }
    }

    public int hashCode() {
        return (block != null ? block.hashCode() : 0);
    }

    public String toString() {
        if( block != null ){
            return block.toString();
        }else{
            return "";
        }
    }
}
