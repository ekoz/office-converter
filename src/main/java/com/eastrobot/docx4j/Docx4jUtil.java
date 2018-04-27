package com.eastrobot.docx4j;

import org.docx4j.Docx4J;
import org.docx4j.Docx4jProperties;
import org.docx4j.convert.out.HTMLSettings;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.OutputStream;

/**
 * Docx4jUtil
 *
 * @author <a href="yogurt_lei@foxmail.com">Yogurt_lei</a>
 * @version v1.0 , 2018-04-26 11:32
 */
public class Docx4jUtil {
    public static void main(String[] args) throws Exception {
        String outpath = "C:\\Users\\User\\Desktop\\kbase-media-2016.docx";
        docToHtml(outpath,"C:\\Users\\User\\Desktop\\test.docx");
        // // 创建包
        // WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
        //
        // wordMLPackage.save(new java.io.File(outpath));
    }

    public static void generate(File inputFile, File outputFile) {}/*{
        InputStream templateStream = null;
        try {
            // Get the template input stream from the application resources.
            final URL resource = inputFile.toURI().toURL();

            // Instanciate the Docx4j objects.
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
            XHTMLImporterImpl XHTMLImporter = new XHTMLImporterImpl(wordMLPackage);

            // Load the XHTML document.
            wordMLPackage.getMainDocumentPart().getContent().addAll(XHTMLImporter.convert(resource));

            // Save it as a DOCX document on disc.
            wordMLPackage.save(outputFile);
            // Desktop.getDesktop().open(outputFile);

        } catch (Exception e) {
            throw new RuntimeException("Error converting file " + inputFile, e);

        } finally {
            if (templateStream != null) {
                try {
                    templateStream.close();
                } catch (Exception ex) {
                    log.error("Can not close the input stream.", ex);

                }
            }
        }
    }*/


    /**
     * docx文档转换为html
     *
     * @param filepath --docx 文件路径f:/1.docx）
     *
     * @return 转换成功返回true, 失败返回false
     */
    public static void docToHtml(String filepath, String outpath) throws Exception {
        boolean bo = false;
        FileWriter fw = null;
            // File infile = new File(filepath);
            // File outfile = new File(outpath);
            // WordprocessingMLPackage wmp = WordprocessingMLPackage.load(infile);
            // // HtmlExporterNonXSLT hn = new HtmlExporterNonXSLT(wmp, new HTMLConversionImageHandler(imgpath, imguri,
            // //         true));
            // String html = (XmlUtils.w3CDomNodeToString(hn.export()));
            // fw = new FileWriter(outfile);
            // fw.write(html);


            WordprocessingMLPackage wordMLPackage= Docx4J.load(new java.io.File(filepath));

            HTMLSettings htmlSettings = Docx4J.createHTMLSettings();
            String imageFilePath = outpath.substring(0, outpath.lastIndexOf("/") + 1) + "/images";
            htmlSettings.setImageDirPath(imageFilePath);
            htmlSettings.setImageTargetUri( "images");
            htmlSettings.setWmlPackage(wordMLPackage);

            String userCSS = "html, body, div, span, h1, h2, h3, h4, h5, h6, p, a, img,  ol, ul, li, table, caption, tbody, tfoot, thead, tr, th, td " +
                    "{ margin: 0; padding: 0; border: 0;}" +
                    "body {line-height: 1;} ";

            htmlSettings.setUserCSS(userCSS);

            OutputStream os = new FileOutputStream(outpath);

            Docx4jProperties.setProperty("docx4j.Convert.Out.HTML.OutputMethodXML", true);

            Docx4J.toHTML(htmlSettings, os, Docx4J.FLAG_EXPORT_PREFER_XSL);
    }
}
