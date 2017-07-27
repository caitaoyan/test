package com.poispace.pojo;

import java.util.List;

public class ExcelModel {

    private String ExcelName;
    private List<CellModel> CellList;

    public String getExcelName() {
        return ExcelName;
    }

    public void setExcelName(String excelName) {
        ExcelName = excelName;
    }

    public List<CellModel> getCellList() {
        return CellList;
    }

    public void setCellList(List<CellModel> cellList) {
        CellList = cellList;
    }

    @Override
    public String toString() {
        return "ExcelModel [ExcelName=" + ExcelName + ", CellList=" + CellList + "]";
    }


}
