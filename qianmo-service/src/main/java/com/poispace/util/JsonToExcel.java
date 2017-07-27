package com.poispace.util;

import com.poispace.pojo.CellModel;
import net.sf.json.JSONArray;
import net.sf.json.JSONObject;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import javax.servlet.http.HttpServletResponse;
import java.awt.Color;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class JsonToExcel {

    /**
     * 解析json，生成excel
     *
     * @param cellList
     * @return
     */
    public void exportExcel(List<CellModel> cellList, HttpServletResponse response) throws Exception {
        // 创建工作簿
        XSSFWorkbook wb = new XSSFWorkbook();
        // 创建工作表Sheet
        Sheet sheet = wb.createSheet();
        DataFormat xssfFormat = wb.createDataFormat();
        Row nRow = null;
        Cell nCell;
        CellRangeAddress region;
        CellModel cellModel;
        String area;
        String areaArray[];//area解析
        String format;//单元格格式
        String formula;//公式
        String content;//单元格内容
        String foregroundColor;//前景色
        XSSFFont font;//字体样式
        XSSFColor xssfColor;//单元格背景色
        Map<String, Object> xssfMap = new HashMap<>();//存放共用格式
        XSSFCellStyle cellStyle;//单元格样式
        int indention;//缩进
        String width;//列宽
        String height;//行高
        String fontSize;//字体大小
        String dataFormat;//格式
        int firstColumn;
        int lastColumn;
        int firstRow;
        int lastRow;
        int oldRow = -1;

        for (int i = 0, cellSize = cellList.size(); i < cellSize; i++) {
            cellModel = cellList.get(i);
            area = cellModel.getArea();
            //System.out.println("area:" + area);
            if (StringUtils.isNotEmpty(area)) {
                areaArray = area.split("_");
                firstColumn = excelColStrToNum(areaArray[0], areaArray[0].length());
                lastColumn = excelColStrToNum(areaArray[2], areaArray[2].length());
                firstRow = Integer.parseInt(areaArray[1]) - 1;
                lastRow = Integer.parseInt(areaArray[3]) - 1;
                if (firstRow != oldRow) {
                    nRow = sheet.createRow(firstRow);
                    oldRow = firstRow;
                }

//                if (!jobj.has("formula")) {
//                    jobj.put("formula", "");
//                }

                // 设置单元格内容
                content = cellModel.getContent();
                nCell = nRow.createCell(firstColumn);

                format = cellModel.getFormat();
                if (StringUtils.isEmpty(content) || "null".equals(content)) {
                    nCell.setCellValue("");
                } else if ("number".equals(format) || "numerical".equals(format)) {
                    double dContent = Double.valueOf(content);
                    nCell.setCellValue(dContent);
                } else if ("percent".equals(format)) {
                    double dContent = Double.valueOf(content.substring(0, content.length() - 1)) / 100;
                    content = String.valueOf(dContent);
                    nCell.setCellValue(dContent);
                } else {
                    nCell.setCellValue(content);
                }

                // 设置单元格的样式
                cellStyle = wb.createCellStyle();

                // 设置字体样式
                font = wb.createFont();// 创建字体对象
                if (StringUtils.isNotEmpty(cellModel.getFont())) {
                    font.setFontName(cellModel.getFont());// 设置字体名称
                    font.setCharSet(XSSFFont.DEFAULT_CHARSET);
                }

                fontSize = cellModel.getFontSize();
                if (StringUtils.isNotEmpty(fontSize)) {
                    font.setFontHeightInPoints((short) Integer.parseInt(fontSize.substring(0, fontSize.indexOf("pt"))));// 设置字体大小
                }

                //加粗
                if ("true".equals(cellModel.getBold())) {
                    font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
                }

                //斜体
                if ("true".equals(cellModel.getItalic())) {
                    font.setItalic(true);
                }

                cellStyle.setFont(font);

                // 设置前景色
                foregroundColor = cellModel.getBackgroundColor();
                //String backgroundColor = jobj.getString("foregroundColor");
                if (xssfMap.containsKey(foregroundColor)) {
                    cellStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
                    cellStyle.setFillForegroundColor((XSSFColor) xssfMap.get(foregroundColor));
                } else if (StringUtils.isNotEmpty(foregroundColor)) {
                    int r2 = Integer.parseInt((foregroundColor.substring(1, 3)), 16);
                    int g2 = Integer.parseInt((foregroundColor.substring(3, 5)), 16);
                    int b2 = Integer.parseInt((foregroundColor.substring(5)), 16);
                    xssfColor = new XSSFColor(new Color(r2, g2, b2));
                    xssfMap.put(foregroundColor, xssfColor);
                    cellStyle.setFillForegroundColor(xssfColor);
                    cellStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
                }

                //设置水平对齐
                switch (cellModel.getAlignment()) {
                    case "center":
                        cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
                        break;
                    case "right":
                        cellStyle.setAlignment(XSSFCellStyle.ALIGN_RIGHT);
                        break;
                    case "left":
                        cellStyle.setAlignment(XSSFCellStyle.ALIGN_LEFT);
                        break;
                }

                //设置垂直对齐
                switch (cellModel.getVertical()) {
                    case "center":
                        cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
                        break;
                    case "bottom":
                        cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_BOTTOM);
                        break;
                    case "top":
                        cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_TOP);
                        break;
                }

                // 设置缩进
                indention = cellModel.getIndentation();
                if (indention != 0) {
                    cellStyle.setIndention((short) indention);
                    cellStyle.setAlignment(CellStyle.ALIGN_LEFT);
                }

                //设置公式
                formula = cellModel.getFormula();

                // 设置单元格格式
                switch (format) {
                    case "text":
                        nCell.setCellType(XSSFCell.CELL_TYPE_STRING);
                        break;
                    case "date":
                        cellStyle.setDataFormat(xssfFormat.getFormat("m/d/yy h:mm"));
                        break;
                    case "numerical":
                        if (StringUtils.isNotEmpty(content) && !"null".equals(content)) {
                            nCell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
                        } else {
                            nCell.setCellType(XSSFCell.CELL_TYPE_BLANK);
                        }
                        if (StringUtils.isNotEmpty(formula) && !"null".equals(formula)) {
                            nCell.setCellFormula(formula.substring(1));
                        }
                        break;
                    case "number":
                        nCell.setCellType(XSSFCell.CELL_TYPE_BLANK);
                        cellStyle.setDataFormat(xssfFormat.getFormat("0.00"));
                        if (StringUtils.isNotEmpty(formula) && !"null".equals(formula)) {
                            nCell.setCellFormula(formula.substring(1));
                        }
                        break;
                    case "boolean":
                        nCell.setCellType(XSSFCell.CELL_TYPE_BOOLEAN);
                        break;
                    case "formula":
                        nCell.setCellType(XSSFCell.CELL_TYPE_FORMULA);
                        break;
                    case "blank":
                        nCell.setCellType(XSSFCell.CELL_TYPE_BLANK);
                        break;
                    case "error":
                        nCell.setCellType(XSSFCell.CELL_TYPE_ERROR);
                        break;
                    case "money":
                        cellStyle.setDataFormat(xssfFormat.getFormat("¥#,##0"));
                        break;
                    case "percent":
                        if (content != null) {
                            int perNum = content.length();
                            int pointIndex = content.indexOf(".");
                            if (pointIndex != -1) {
                                int totalNum = perNum - pointIndex - 3;
                                String perFormat = "0";
                                for (int j = 0; j < totalNum; j++) {
                                    if (j == 0) {
                                        perFormat += ".";
                                    }
                                    perFormat += "0";
                                }
                                cellStyle.setDataFormat(xssfFormat.getFormat(perFormat + "%"));
                            } else {
                                cellStyle.setDataFormat(xssfFormat.getFormat("0%"));
                                nCell.setCellType(XSSFCell.CELL_TYPE_BLANK);
                            }
                        } else {
                            nCell.setCellType(XSSFCell.CELL_TYPE_BLANK);
                        }
                        break;
                    case "notation":
                        cellStyle.setDataFormat(xssfFormat.getFormat("0.00E+00"));
                        break;
                    case "fraction":
                        cellStyle.setDataFormat(xssfFormat.getFormat("??/??"));
                        break;
                    default:
                        break;
                }


                //设置单元格格式
//                dataFormat = jobj.getString("dataFormat");
//                if (dataFormat != null) {
//                    cellStyle.setDataFormat(Integer.parseInt(dataFormat));
//                }

                // 保存合并单元格数据
                if (firstColumn != lastColumn || firstRow != lastRow) {
                    region = new CellRangeAddress(firstRow, lastRow, firstColumn, lastColumn);
                    sheet.addMergedRegion(region);
                }

                //设置列宽
                width = cellModel.getWidth();
                if (StringUtils.isNotEmpty(width)) {
                    sheet.setColumnWidth(firstColumn, (int) (Math.round(Double.parseDouble(width) * 36.56)));
                }

                //设置行高
                height = cellModel.getHeight();
                if (StringUtils.isNotEmpty(height)) {
                    nRow.setHeightInPoints(Float.parseFloat(height));
                }

                //设置自动换行
                if ("true".equals(cellModel.getWrapText())) {
                    cellStyle.setWrapText(true);
                }

                // 添加左边框
                if ("true".equals(cellModel.getLeftFrame())) {
                    cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
                }
                // 添加右边框
                if ("true".equals(cellModel.getRightFrame())) {
                    cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
                }
                // 添加上边框
                if ("true".equals(cellModel.getTopFrame())) {
                    cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
                }
                // 添加下边框
                if ("true".equals(cellModel.getBottomFrame())) {
                    cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
                }

                nCell.setCellStyle(cellStyle);
            }
        }

        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        wb.write(byteArrayOutputStream);
        // 工具类，封装弹出下载框：
        Date now = new Date();
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMddHH24mmss");
        String nowTime = dateFormat.format(now);
        String outFile = "report-" + nowTime + ".xlsx";
        DownloadUtil down = new DownloadUtil();
        down.download(byteArrayOutputStream, response, outFile);
        byteArrayOutputStream.close();
    }

    public String getChangeContent(JSONArray jsonArray) {
        //System.out.println("jsonArray.size():"+jsonArray.size());
        for (int i = 0; i < jsonArray.size(); i++) {
            JSONObject node12 = JSONObject.fromObject(jsonArray.getJSONObject(i));
            //JSONArray jsonArray12 = JSONArray.fromObject(node12);
            //System.out.println(jsonArray12.toString());
            Map<String, String> map = new HashMap<String, String>();

            String coord = (String) node12.get("coord");
            String value = (String) node12.get("value");
            value = changeContent(value);

            map.put("coord", coord);
            map.put("value", value);

            jsonArray.set(i, JSONObject.fromObject(map));
        }
        return jsonArray.toString();
    }

    public String changeContent(String value) {
        StringBuilder strBuilder = new StringBuilder(value);
        for (int i = 0; i < strBuilder.length(); i++) {
            char c = strBuilder.charAt(i);
            int num = (int) c;
            if ((num >= 65 && num <= 90) || (num >= 97 && num <= 122)) {
                strBuilder.setCharAt(i, '1');
            }
        }
        return strBuilder.toString();
    }


    /**
     * Excel column index begin 1
     *
     * @param colStr
     * @param length
     * @return
     */
    public static int excelColStrToNum(String colStr, int length) {
        int num;
        int result = 0;
        for (int i = 0; i < length; i++) {
            char ch = colStr.charAt(length - i - 1);
            num = ch - 'A' + 1;
            num *= Math.pow(26, i);
            result += num;
        }
        return result - 1;
    }

}
