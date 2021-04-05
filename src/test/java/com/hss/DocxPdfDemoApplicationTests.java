package com.hss;


import com.hss.util.Word2PdfUtils;
import com.hss.util.WordTemplate2PdfUtil;
import com.hss.util.WordTemplateUtil;
import lombok.extern.slf4j.Slf4j;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.FileNotFoundException;
import java.util.HashMap;

@RunWith(SpringRunner.class)
@SpringBootTest
@Slf4j
public class DocxPdfDemoApplicationTests {

    @Value("${spring.servlet.multipart.location}")
    private String location;

    @Test
    public void contextLoads() {

    }

    /**
     * word文档转pdf
     */
    /*@Test
    public void word2Pdf(){
        String docFileName = "test0001委托担保申请书.docx";

        String pdfFileName = "test0001委托担保申请书.pdf";

        String docxPath = location + docFileName;

        String pdfPath = location + pdfFileName;

        try {
            Word2PdfUtils.convertDocxToPdf(docxPath,pdfPath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }*/

    /**
     * 根据world模板，生成world文件，在将world文件转成pdf
     * 占位符：${xxx}，支持驼峰命名。
     * 遇到关键字不生效时，替换占位符名称即可。
     */
    /*@Test
    public void worldTemplate(){
        //字典项信息
        String orderNo = "test0001";
        String areaAddress = "黄山市";
        String startTime = "2021年04月05日";
        String tenrDay = "180";
        String projectName = "项目名称";

        //文件信息
        String docFileName = "委托担保申请书.docx";

        String docxPath = location + docFileName;

        String newFileName = orderNo +"委托担保申请书";
        String newDocxPath = location + newFileName + ".docx";
        String newPdfPath = location + newFileName + ".pdf";

        HashMap<String,String> data = new HashMap<>();
        data.put("orderNoField",orderNo);
        data.put("areaAddressField",areaAddress);
        data.put("startTimeField",startTime);
        data.put("tenrDayField",tenrDay);
        data.put("projectNameField",projectName);

        try {
            //由模板生成word文件
            WordTemplateUtil.replaceData(docxPath,newDocxPath,data);
            //将word文件转为pdf
            Word2PdfUtils.convertDocxToPdf(newDocxPath,newPdfPath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }*/

    /**
     * 根据world模板生成pdf文件
     * 占位符：{xxx},不支持驼峰标识。全小写
     * 优点：由模板直接生成pdf文件，一步到位
     */
    /*@Test
    public void worldTemplate2Pdf(){
        //字典项信息
        String orderNo = "test0001";
        String areaAddress = "黄山市";
        String startTime = "2021年04月05日";
        String tenrDay = "180";
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

    }*/
}
