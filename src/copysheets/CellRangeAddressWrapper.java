package copysheets;

import org.apache.poi.ss.util.CellRangeAddress;

/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author Guillermo
 */

class CellRangeAddressWrapper implements Comparable<copysheets.CellRangeAddressWrapper> {  

public CellRangeAddress range;  

/** 
 * @param theRange the CellRangeAddress object to wrap. 
 */  
public CellRangeAddressWrapper(CellRangeAddress theRange) {  
      this.range = theRange;  
}  

/** 
 * @param o the object to compare. 
 * @return -1 the current instance is prior to the object in parameter, 0: equal, 1: after... 
 */  
public int compareTo(copysheets.CellRangeAddressWrapper o) {  

            if (range.getFirstColumn() < o.range.getFirstColumn()  
                        && range.getFirstRow() < o.range.getFirstRow()) {  
                  return -1;  
            } else if (range.getFirstColumn() == o.range.getFirstColumn()  
                        && range.getFirstRow() == o.range.getFirstRow()) {  
                  return 0;  
            } else {  
                  return 1;  
            }  

}  
}
