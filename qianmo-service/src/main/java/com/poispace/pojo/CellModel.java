package com.poispace.pojo;

public class CellModel {
    private String cellName;
    private String content;
    private String format;
    private String area;
    private String font;
    private String fontSize;
    private String foregroundColor;
    private String backgroundColor;
    private String formula;
    private String leftFrame;
    private String topFrame;
    private String rightFrame;
    private String bottomFrame;
    private short indentation;
    private String alignment;
    private String bold;
    private String italic;
    private String wrapText;
    private String width;
    private String height;
    private String vertical;
    private int dataFormat;

    public CellModel() {
        // TODO Auto-generated constructor stub
    }

    public int getDataFormat() {
        return dataFormat;
    }

    public void setDataFormat(int dataDormat) {
        this.dataFormat = dataDormat;
    }

    public String getVertical() {
        return vertical;
    }

    public void setVertical(String vertical) {
        this.vertical = vertical;
    }

    public String getWidth() {
        return width;
    }

    public void setWidth(String width) {
        this.width = width;
    }

    public String getHeight() {
        return height;
    }

    public void setHeight(String height) {
        this.height = height;
    }

    public String getWrapText() {
        return wrapText;
    }

    public void setWrapText(String wrapText) {
        this.wrapText = wrapText;
    }

    public String getCellName() {
        return cellName;
    }

    public void setCellName(String cellName) {
        this.cellName = cellName;
    }

    public String getContent() {
        return content;
    }

    public void setContent(String content) {
        this.content = content;
    }

    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }

    public String getArea() {
        return area;
    }

    public void setArea(String area) {
        this.area = area;
    }

    public String getFont() {
        return font;
    }

    public void setFont(String font) {
        this.font = font;
    }


    public String getFormula() {
        return formula;
    }

    public void setFormula(String formula) {
        this.formula = formula;
    }

    public String getFontSize() {
        return fontSize;
    }

    public void setFontSize(String fontSize) {
        this.fontSize = fontSize;
    }

    public String getForegroundColor() {
        return foregroundColor;
    }

    public void setForegroundColor(String foregroundColor) {
        this.foregroundColor = foregroundColor;
    }

    public String getBackgroundColor() {
        return backgroundColor;
    }

    public void setBackgroundColor(String backgroundColor) {
        this.backgroundColor = backgroundColor;
    }


    public String getLeftFrame() {
        return leftFrame;
    }

    public void setLeftFrame(String leftFrame) {
        this.leftFrame = leftFrame;
    }

    public String getTopFrame() {
        return topFrame;
    }

    public void setTopFrame(String topFrame) {
        this.topFrame = topFrame;
    }

    public String getRightFrame() {
        return rightFrame;
    }

    public void setRightFrame(String rightFrame) {
        this.rightFrame = rightFrame;
    }

    public String getBottomFrame() {
        return bottomFrame;
    }

    public void setBottomFrame(String bottomFrame) {
        this.bottomFrame = bottomFrame;
    }


    public short getIndentation() {
        return indentation;
    }

    public void setIndentation(short indentation) {
        this.indentation = indentation;
    }


    public String getAlignment() {
        return alignment;
    }

    public void setAlignment(String alignment) {
        this.alignment = alignment;
    }

    public String getBold() {
        return bold;
    }

    public void setBold(String bold) {
        this.bold = bold;
    }

    public String getItalic() {
        return italic;
    }

    public void setItalic(String italic) {
        this.italic = italic;
    }


    public boolean allEmpty() {

        if (cellName == null || cellName.trim().equals("")) {
            if (content == null || content.trim().equals("")) {
                if (format == null || format.trim().equals("")) {
                    if (area == null || area.trim().equals("")) {
                        if (font == null || font.trim().equals("")) {
                            if (foregroundColor == null || foregroundColor.trim().equals("")) {
                                if (backgroundColor == null || backgroundColor.trim().equals("")) {


                                    if (fontSize == null || fontSize.equals("0pt")) {
                                        if (indentation == 0) {
                                            return true;
                                        }


                                    }

                                }

                            }

                        }
                    }
                }
            }

        }

        return false;

    }

}
