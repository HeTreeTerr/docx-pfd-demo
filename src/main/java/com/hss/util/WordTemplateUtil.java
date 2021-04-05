package com.hss.util;

import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

import java.beans.PropertyDescriptor;
import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * word根据模板生成文件
 */
public class WordTemplateUtil {

    /**
     * 加载模板并替换数据
     *
     * @param templatePath
     * @param outPutPath
     * @param data
     * @throws Exception
     */
    public static void replaceData(String templatePath, String outPutPath, HashMap<String, String> data) throws Exception {
        final String TEMPLATE_NAME = templatePath;
        InputStream templateInputStream = new FileInputStream(TEMPLATE_NAME);
        //加载模板文件并创建WordprocessingMLPackage对象
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(templateInputStream);
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
        VariablePrepare.prepare(wordMLPackage);
        documentPart.variableReplace(data);
        OutputStream os = new FileOutputStream(new File(outPutPath));
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        wordMLPackage.save(outputStream);
        outputStream.writeTo(os);
        os.close();
        outputStream.close();
        templateInputStream.close();
    }

    /**
     * 如果从数据库查出来的是不同对象集合，那么该方法直接把对象的属性和值直接存为Map
     * 比如示例中可以是两个对象，ClassInfo(className,total,male,female)、student(name,sex,...)
     * 前提是模板中的占位符一定要和对象的属性名一一对应
     * @param objlist
     * @return
     * @throws Exception
     */
    public HashMap<String, String> toMap(List<Object> objlist) throws Exception {
        HashMap<String, String> variables = new HashMap<>();
        String value = "";
        if (objlist == null && objlist.isEmpty()) {
            return null;
        } else {
            for (Object obj : objlist) {
                Field[] fields = null;
                fields = obj.getClass().getDeclaredFields();
                //删除字段数组里的serialVersionUID
                for (int i = 0; i < fields.length; i++) {
                    fields[i].setAccessible(true);
                    if (fields[i].getName().equals("serialVersionUID")) {
                        fields[i] = fields[fields.length - 1];
                        fields = Arrays.copyOf(fields, fields.length - 1);
                        break;
                    }
                }
                //遍历数组，获取属性名及属性值
                for (Field field : fields) {
                    field.setAccessible(true);
                    String fieldName = field.getName();
                    //如果属性类型是Date
                    if (field.getGenericType().toString().equals("class java.util.Date")) {
                        PropertyDescriptor pd = new PropertyDescriptor(fieldName, obj.getClass());
                        Method getMethod = pd.getReadMethod();
                        Date date = (Date) getMethod.invoke(obj);
                        value = (date == null) ? "" : new SimpleDateFormat("yyyy-MM-dd").format(date);
                    } else if (field.getGenericType().toString().equals("class java.lang.Integer")) {
                        //如果属性类型是Integer
                        PropertyDescriptor pd = new PropertyDescriptor(fieldName, obj.getClass());
                        Method getMethod = pd.getReadMethod();
                        Object object = getMethod.invoke(obj);
                        value = (object == null) ? String.valueOf(0) : object.toString();

                    } else if (field.getGenericType().toString().equals("class java.lang.Double")) {
                        //如果属性类型是Double
                        PropertyDescriptor pd = new PropertyDescriptor(fieldName, obj.getClass());
                        Method getMethod = pd.getReadMethod();
                        Object object = getMethod.invoke(obj);
                        value = (object == null) ? String.valueOf(0.0) : object.toString();

                    } else {
                        PropertyDescriptor pd = new PropertyDescriptor(fieldName, obj.getClass());
                        Method getMethod = pd.getReadMethod();
                        Object object = getMethod.invoke(obj);
                        value = (object == null) ? "  " : object.toString();

                    }
                    variables.put(fieldName, value);
                }
            }
            return variables;
        }
    }

}
