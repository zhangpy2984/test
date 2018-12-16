package poi;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

/**
 * 自定义row
 *
 * @author zhangpengyue
 * @date 2018/12/16
 */
public class XRow implements Serializable {
    private static final long serialVersionUID = 490011564671029222L;
    //行号
    private int rowNum;

    private List<XCell> cellList;

    public XRow(int rowNum) {
        this.rowNum = rowNum;
        cellList = new ArrayList<>();
    }

    public int getRowNum() {
        return rowNum;
    }

    public void setRowNum(int rowNum) {
        this.rowNum = rowNum;
    }

    public List<XCell> getCellList() {
        return cellList;
    }

    public void setCellList(List<XCell> cellList) {
        this.cellList = cellList;
    }

    public synchronized void addCell(XCell cell) {
        cellList.add(cell);
    }
}
