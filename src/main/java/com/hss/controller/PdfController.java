package com.hss.controller;

import com.hss.util.WordTemplate2PdfUtil;
import lombok.extern.slf4j.Slf4j;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import java.io.FileNotFoundException;
import java.util.HashMap;

@RestController
@Slf4j
public class PdfController {

    @Value("${spring.servlet.multipart.location}")
    private String location;

    @RequestMapping(value = "/worldTemplate2Pdf")
    public String worldTemplate2Pdf(@RequestParam(value = "tenrDay",defaultValue = "180") String tenrDay){
        //字典项信息
        String orderNo = "test0001";
        String areaAddress = "黄山市";
        String startTime = "2021年04月05日";
        String projectName = "项目名称";

        //文件信息
        String docFileName = "委托担保申请书.docx";

        String docxPath = location + docFileName;

        String newFileName = orderNo +"委托担保申请书";
        String newPdfPath = location + newFileName + ".pdf";

        log.info(docxPath);
        log.info(newPdfPath);
        try {
            //获取模板文件对象
            WordprocessingMLPackage template = WordTemplate2PdfUtil.getTemplate(docxPath);
            HashMap<String,String> data = new HashMap<>();
            data.put("{name}","张三");
            data.put("{ordernofield}",orderNo);
            data.put("{areaaddressfield}",areaAddress);
            data.put("{starttimefield}",startTime);
            data.put("{tenrdayfield}",tenrDay);
            data.put("{projectnamefield}",projectName);
            //替换模板中文件内容
            WordTemplate2PdfUtil.replacePlaceholder(template, data);
            //生成pdf
            WordTemplate2PdfUtil.writeDocxToStream(template, newPdfPath);
        } catch (Docx4JException e) {
            e.printStackTrace();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return "success";
    }
}
