package net.sf.jxls;

import junit.framework.TestCase;
import net.sf.jxls.transformation.ShiftTransformation;
import net.sf.jxls.tag.Block;

/**
 * @author Leonid Vysochyn
 */
public class ShiftTransformationTest extends TestCase {
    public void testEqualsTrue(){
        Block b1 = new Block(1, (short)2, 3, (short) 4);
        Block b2 = new Block(1, (short)2, 3, (short) 4);
        ShiftTransformation st1 = new ShiftTransformation(b1, 1, 2);
        ShiftTransformation st2 = new ShiftTransformation(b2, 1, 2);
        assertTrue (st1.equals(st2));
        st1 = new ShiftTransformation(null, 2, 3);
        st2 = new ShiftTransformation(null, 2, 3);
        assertTrue( st1.equals(st2) );
    }

    public void testEqualsFalse(){
        Block b1 = new Block(1, (short)2, 3, (short) 4);
        Block b2 = new Block(1, (short)2, 3, (short) 5);
        ShiftTransformation st1 = new ShiftTransformation(b1, 1, 2);
        ShiftTransformation st2 = new ShiftTransformation(b2, 1, 2);
        assertFalse (st1.equals(st2));
        st2 = new ShiftTransformation(b1, 0, 2);
        assertFalse (st1.equals(st2));
        st2 = new ShiftTransformation(b1, 1, 3);
        assertFalse (st1.equals(st2));
        st1 = new ShiftTransformation(null, 1, 3);
        assertFalse( st1.equals(st2) );
    }
}
