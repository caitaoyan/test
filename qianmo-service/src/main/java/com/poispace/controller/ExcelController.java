package com.poispace.controller;

import com.poispace.pojo.CellModel;
import com.poispace.util.JsonToExcel;
import com.poispace.util.PoiToJson;
import net.sf.json.JSONArray;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.List;

/**
 * 文件上传解析
 *
 * @author zhengding @version1.0
 */
@Controller
public class ExcelController {

    @RequestMapping(value = "/excel", method = {RequestMethod.POST, RequestMethod.GET})
    public String showExcel() {
        System.out.println("showExcel");
        return "redirect:app.html?url=" + "111.txt";
    }

    /**
     * 获取excel，解析生成json
     *
     * @param file
     * @param request
     * @param model
     * @return
     */
    @RequestMapping(value = "/excelPoi", method = {RequestMethod.POST, RequestMethod.GET})
    public String excelPoi(MultipartFile file, HttpServletRequest request, Model model) {
        //System.err.println("是否为中文1111");
        String newjsonfileName = PoiToJson.excelTojson(file, request);
        model.addAttribute("url", newjsonfileName);
        //return "qianmo";
        return "redirect:app.html?url=" + newjsonfileName;
    }

    /**
     * 获取文档中的json数据
     *
     * @param request
     * @param response
     * @param url
     */
    @RequestMapping(value = "/getContentJson", method = {RequestMethod.POST, RequestMethod.GET})
    @ResponseBody
    public void getContentJson(HttpServletRequest request, HttpServletResponse response, @RequestBody String url) {
        PoiToJson.getContent(request, response, url);
    }

    /**
     * 解析json，生成excel，弹出下载框
     *
     * @param response
     * @param cellList
     * @throws Exception
     */
    @RequestMapping(value = "/excelDownload", method = {RequestMethod.POST, RequestMethod.GET})
    public void excelDownload(HttpServletResponse response, @RequestBody List<CellModel> cellList) throws Exception {
        JsonToExcel jsonToExcel = new JsonToExcel();
        jsonToExcel.exportExcel(cellList, response);
        //SimpleDateFormat sdf = new SimpleDateFormat("YYYY/MM/dd HH:mm:ss.SSS");
        //JSONArray celljson = JSONArray.fromObject(cellList);
        //System.out.println(sdf.format(new Date()));
        //System.out.println(sdf.format(new Date()));
    }

    /**
     * 单元格公式转换
     *
     * @param
     * @return
     */
    @RequestMapping(value = "/changeContent", method = {RequestMethod.POST, RequestMethod.GET})
    @ResponseBody
    public String changeContent(HttpServletResponse response, @RequestBody String content) {
        JSONArray jsonArray = JSONArray.fromObject(content);
        String newJsonArray = new JsonToExcel().getChangeContent(jsonArray);
        response.setHeader("Access-Control-Allow-Origin", "*");
        return newJsonArray;
    }
}
