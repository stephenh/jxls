package net.sf.jxls.tag;

import net.sf.jxls.transformer.Sheet;

/**
 * Represents rectangular range of excel cells
 * @author Leonid Vysochyn
 */
public class Block {
    int startRowNum;
    int endRowNum;
    short startCellNum;
    short endCellNum;

    Sheet sheet;

    public Block(Sheet sheet, int startRowNum, int endRowNum) {
        this.startRowNum = startRowNum;
        this.endRowNum = endRowNum;
        this.startCellNum = -1;
        this.endCellNum = -1;
        this.sheet = sheet;
    }

    public Block(int startRowNum, short startCellNum, int endRowNum, short endCellNum) {
        this.startRowNum = startRowNum;
        this.startCellNum = startCellNum;
        this.endRowNum = endRowNum;
        this.endCellNum = endCellNum;
    }

    public Block horizontalShift(short cellShift){
        startCellNum += cellShift;
        endCellNum += cellShift;
        return this;
    }

    public Block verticalShift(int rowShift){
        startRowNum += rowShift;
        endRowNum += rowShift;
        return this;
    }

    public short getStartCellNum() {
        return startCellNum;
    }

    public void setStartCellNum(short startCellNum) {
        this.startCellNum = startCellNum;
    }

    public short getEndCellNum() {
        return endCellNum;
    }

    public void setEndCellNum(short endCellNum) {
        this.endCellNum = endCellNum;
    }

    public int getStartRowNum() {
        return startRowNum;
    }

    public void setStartRowNum(int startRowNum) {
        this.startRowNum = startRowNum;
    }

    public int getEndRowNum() {
        return endRowNum;
    }

    public void setEndRowNum(int endRowNum) {
        this.endRowNum = endRowNum;
    }

    public int getNumberOfRows(){
        return endRowNum - startRowNum + 1;
    }

    public int getNumberOfColumns(){
        return endCellNum - startCellNum + 1;
    }

    public boolean contains(int rowNum, int cellNum){
        return (startRowNum <= rowNum && rowNum <= endRowNum && ((startCellNum==-1 && endCellNum==-1) || (startCellNum <= cellNum && cellNum <= endCellNum)));
    }

    public boolean contains(Point p){
        return (startRowNum <= p.getRow() && p.getRow() <= endRowNum &&
                ((startCellNum<0 || endCellNum<0) || (startCellNum <= p.getCol() && p.getCol() <= endCellNum)));
    }

    public boolean isAbove(Point p){
        return (endRowNum < p.getRow());
    }

    public boolean isBelow(Point p){
        return (startRowNum > p.getRow());
    }

    public boolean isRowBlock(){
        return (startCellNum < 0 || endCellNum < 0 || (startCellNum > endCellNum) );
    }

    public boolean isColBlock(){
        return (startRowNum <0 || endRowNum < 0 || (startRowNum > endRowNum) );
    }

    public Sheet getSheet() {
        return sheet;
    }

    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;

        final Block block = (Block) o;

        if (endCellNum != block.endCellNum) return false;
        if (endRowNum != block.endRowNum) return false;
        if (startCellNum != block.startCellNum) return false;
        if (startRowNum != block.startRowNum) return false;
        if (sheet != null ? !sheet.equals(block.sheet) : block.sheet != null) return false;

        return true;
    }

    public int hashCode() {
        int result;
        result = startRowNum;
        result = 29 * result + endRowNum;
        result = 29 * result + (int) startCellNum;
        result = 29 * result + (int) endCellNum;
        result = 29 * result + (sheet != null ? sheet.hashCode() : 0);
        return result;
    }

    public String toString() {
        return "Block (" + startRowNum + ", " + startCellNum + ", " + endRowNum + ", " + endCellNum + ")";
    }
}
