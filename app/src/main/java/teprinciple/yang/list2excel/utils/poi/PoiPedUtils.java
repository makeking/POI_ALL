package teprinciple.yang.list2excel.utils.poi;


import android.content.res.AssetManager;
import android.graphics.Bitmap;
import android.graphics.BitmapFactory;
import android.graphics.Canvas;
import android.graphics.Paint;
import android.graphics.pdf.PdfDocument;
import android.util.Log;

import com.tom_roush.pdfbox.pdmodel.PDDocument;
import com.tom_roush.pdfbox.pdmodel.PDPage;
import com.tom_roush.pdfbox.pdmodel.PDPageContentStream;
import com.tom_roush.pdfbox.pdmodel.font.PDFont;
import com.tom_roush.pdfbox.pdmodel.font.PDType1Font;
import com.tom_roush.pdfbox.pdmodel.graphics.color.PDColor;
import com.tom_roush.pdfbox.pdmodel.graphics.image.JPEGFactory;
import com.tom_roush.pdfbox.pdmodel.graphics.image.LosslessFactory;
import com.tom_roush.pdfbox.pdmodel.graphics.image.PDImageXObject;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * 使用poi 导出pdf 的工具类
 */
public class PoiPedUtils {
    /**
     * 1. 生成doc 样式的内容
     *
     * @param path
     */
//    public static void exportDoc(String path) {
//        try {
//        PDEmbeddedFilesNameTreeNode efTree = new PDEmbeddedFilesNameTreeNode();
//        //first create the file specification, which holds the embedded file
//        PDComplexFileSpecification fs = new PDComplexFileSpecification();
//        fs.setFile( path );
//        InputStream is = null;
////            is = new FileInputStream(new File(path));
//        PDEmbeddedFile ef = new PDEmbeddedFile(doc, is );
//        //set some of the attributes of the embedded file
//        ef.setSubtype( "test/plain" );
//        ef.setSize( data.length );
//        ef.setCreationDate( new GregorianCalendar() );
//        fs.setEmbeddedFile( ef );
//
//        //now add the entry to the embedded file tree and set in the document.
//        Map efMap = new HashMap();
//        efMap.put( "My first attachment", fs );
//        efTree.setNames( efMap );
//        //attachments are stored as part of the "names" dictionary in the document catalog
//        PDDocumentNameDictionary names = new PDDocumentNameDictionary( doc.getDocumentCatalog() );
//        names.setEmbeddedFiles( efTree );
//        doc.getDocumentCatalog().setNames( names );
//        } catch (FileNotFoundException e) {
//            e.printStackTrace();
//        }
//    }
    public static void exportDoc(String path) {

    }

    //    public static void exportPdf(String fileName) {
//        PdfDocument pdfDocument = new PdfDocument();
//        PdfDocument.PageInfo pageInfo = new PdfDocument.PageInfo.Builder(100, 100, 1).create();
//        PdfDocument.Page page = pdfDocument.startPage(pageInfo);
//        Canvas canvas = page.getCanvas();
//        //设置Canvas大小，必须跟pdf大小相同
//        canvas.scale(0.35f,0.35f);
//        canvas.drawText("测试，这是一行文字",0,100,new Paint());
//        pdfDocument.finishPage(page);
//        try {
//            pdfDocument.writeTo(new FileOutputStream(new File(fileName)));
//            System.out.println("文件生成成功！");
//        } catch (IOException e) {
//            e.printStackTrace();
//            System.out.println("文件生成失败！");
//        }
//
//
//    }

//    public static void exportPdf(AssetManager assetManager, String fileName) {
//        PDDocument document = new PDDocument();
//        PDPage page = new PDPage();
//        document.addPage(page);
//        PDFont font = PDType1Font.HELVETICA;
//        PDPageContentStream contentStream;
//
//        try {
//            // Define a content stream for adding to the PDF
//            contentStream = new PDPageContentStream(document, page);
//
//            // Write Hello World in blue text
//            contentStream.beginText();
//            contentStream.setNonStrokingColor(15, 38, 192);
//            contentStream.setFont(font, 12);
//            contentStream.newLineAtOffset(100, 700);
//            contentStream.showText("Hello World");
//            contentStream.endText();
//
//            // Load in the images
//            InputStream in = assetManager.open("falcon.jpg");
//            InputStream alpha = assetManager.open("trans.png");
//
//            // Draw a green rectangle
//            contentStream.addRect(5, 500, 100, 100);
//            contentStream.setNonStrokingColor(0, 255, 125);
//            contentStream.fill();
//
//            // Draw the falcon base image
//            PDImageXObject ximage = JPEGFactory.createFromStream(document, in);
//            contentStream.drawImage(ximage, 20, 20);
//
//            // Draw the red overlay image
//            Bitmap alphaImage = BitmapFactory.decodeStream(alpha);
//            PDImageXObject alphaXimage = LosslessFactory.createFromImage(document, alphaImage);
//            contentStream.drawImage(alphaXimage, 20, 20);
//
//            // Make sure that the content stream is closed:
//            contentStream.close();
//
//            // Save the final pdf document to a file
//            document.save(fileName);
//            document.close();
//        } catch (IOException e) {
//        }
//    }
//    // 1. 设置标题
//    public static void exportTitlePdf(String fileName, String title, int size,int r,int g,int b,int tx,int ty) {
//        PDDocument document = new PDDocument();
//
//        PDPage page = new PDPage();
//        document.addPage(page);
//        PDFont font = PDType1Font.HELVETICA;
//        PDPageContentStream contentStream;
//        try {
//            // Define a content stream for adding to the PDF
//            contentStream = new PDPageContentStream(document, page);
//            contentStream.beginText();
//            contentStream.setStrokingColor(r,g,b);
//            contentStream.setFont(font, size);
//            contentStream.newLineAtOffset(tx, ty);
//            contentStream.showText(title);
//            contentStream.endText();
//            contentStream.close();
//            // Save the final pdf document to a file
//            document.save(fileName);
//            document.close();
//        } catch (IOException e) {
//        }
//    }


}


