package com.yuanzi.office2pdf;

import java.io.File;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

/**
 * @Description:jacob + msOfficeWord + SaveAsPDFandXPS 
 * 				（完美保持原doc格式，效率最慢，只能在windows环境下进行）
 * 				而且速度是最慢的需要安装msofficeWord以及SaveAsPDFandXPS.exe
 * 				(word的一个插件，用来把word转化为pdf)
 * 				有Office软件也行,自带上面插件
 * @author xueyya
 * @date:2018年3月15日 下午3:35:08
 */
public class Jacob {
	
	/** 转PDF格式值 */
	static final int WORD_FORMAT_PDF = 17;
	static final int EXCEL_FORMAT_PDF = 0;
	static final int PPT_FORMAT_PDF = 32;
	
	/**
	 * @Description:根据文件类型转换为pdf
	 * @author xueyya
	 * @date:2018年3月16日 上午9:16:47
	 * @param inputFile
	 * @param pdfFile void
	 */
	public static void convert2PDF(String inputFile, String pdfFile) {
		String suffix = getFileSufix(inputFile);
		if (suffix.equals("doc") || suffix.equals("docx") || suffix.equals("txt")) {
	        word2PDF(inputFile, pdfFile);
	    } else if (suffix.equals("xls") || suffix.equals("xlsx")) {
	    	excel2PDF(inputFile, pdfFile);
	    } else if (suffix.equals("ppt") || suffix.equals("pptx")) {
	        ppt2PDF(inputFile, pdfFile);
	    } else {
	        System.out.println("文件格式不支持转换!");
	    }
	}
	
	/**
	 * @Description:word转pdf
	 * @author xueyya
	 * @date:2018年3月15日 下午4:07:49
	 * @param inputFile void
	 * @param pdfFile 
	 */
	private static void word2PDF(String inputFile, String pdfFile) {    
        System.out.println("启动Word...");      
        long start = System.currentTimeMillis();      
        ActiveXComponent app = null;  
        Dispatch doc = null;  
        try {      
        	// 创建一个word对象
            app = new ActiveXComponent("Word.Application");      
            // 不可见打开word
            app.setProperty("Visible", new Variant(false));  
            // 获取文挡属性
            Dispatch docs = app.getProperty("Documents").toDispatch();    
            // 调用Documents对象中Open方法打开文档，并返回打开的文档对象Document
            doc = Dispatch.call(docs, "Open", inputFile).toDispatch();  
            System.out.println("打开文档..." + inputFile);  
            System.out.println("转换文档到PDF..." + pdfFile);      
            File tofile = new File(pdfFile);      
            if(tofile.exists()) {      
                tofile.delete();      
            }      
            // word保存为pdf格式宏，值为17
            Dispatch.call(doc, "SaveAs", pdfFile, WORD_FORMAT_PDF);      
            long end = System.currentTimeMillis();      
            System.out.println("转换完成..用时：" + (end - start) + "ms.");  
        } catch (Exception e) {      
            System.out.println("========Error:文档转换失败：" + e.getMessage());      
        } finally {  
            Dispatch.call(doc, "Close", false);  
            System.out.println("关闭文档");  
            if (app != null)      
                app.invoke("Quit", new Variant[] {});      
            }  
          //如果没有这句话,winword.exe进程将不会关闭  
           ComThread.Release();
    }
	
	/**
	 * @Description:excel转pdf
	 * @author xueyya
	 * @date:2018年3月15日 下午4:07:49
	 * @param inputFile void
	 * @param pdfFile 
	 */
	private static void excel2PDF(String inputFile, String pdfFile) {    
        System.out.println("启动Excel...");      
        long start = System.currentTimeMillis();      
        ActiveXComponent app = null;  
        Dispatch excel = null;  
        try {      
        	// 创建一个excel对象
            app = new ActiveXComponent("Excel.Application");      
            // 不可见打开excel
            app.setProperty("Visible", new Variant(false));  
            // 获取文挡属性
            Dispatch excels = app.getProperty("Workbooks").toDispatch();    
            // 调用Documents对象中Open方法打开文档，并返回打开的文档对象Document
            excel = Dispatch.call(excels, "Open", inputFile).toDispatch();  
            System.out.println("打开文档..." + inputFile);  
            System.out.println("转换文档到PDF..." + pdfFile);      
            File tofile = new File(pdfFile);      
            if(tofile.exists()) {      
                tofile.delete();      
            }      
            // Excel不能调用SaveAs方法
            Dispatch.call(excel, "ExportAsFixedFormat", EXCEL_FORMAT_PDF, pdfFile);
            long end = System.currentTimeMillis();      
            System.out.println("转换完成..用时：" + (end - start) + "ms.");  
        } catch (Exception e) {      
            System.out.println("========Error:文档转换失败：" + e.getMessage());      
        } finally {  
            Dispatch.call(excel, "Close", false);  
            System.out.println("关闭文档");  
            if (app != null)      
                app.invoke("Quit", new Variant[] {});      
            }  
          //如果没有这句话,winword.exe进程将不会关闭  
           ComThread.Release();
    }
	
	/**
	 * @Description:ppt转pdf
	 * @author xueyya
	 * @date:2018年3月15日 下午4:07:49
	 * @param inputFile void
	 * @param pdfFile 
	 */
	private static void ppt2PDF(String inputFile, String pdfFile) {    
        System.out.println("启动PPT...");      
        long start = System.currentTimeMillis();      
        ActiveXComponent app = null;  
        Dispatch ppt = null;  
        try {      
        	// 创建一个ppt对象
            app = new ActiveXComponent("PowerPoint.Application");      
            // 不可见打开（PPT转换不运行隐藏，所以这里要注释掉）
            // app.setProperty("Visible", new Variant(false));  
            // 获取文挡属性
            Dispatch ppts = app.getProperty("Presentations").toDispatch();    
            // 调用Documents对象中Open方法打开文档，并返回打开的文档对象Document
            ppt = Dispatch.call(ppts, "Open", inputFile, true, true, false).toDispatch();  
            System.out.println("打开文档..." + inputFile);  
            System.out.println("转换文档到PDF..." + pdfFile);      
            File tofile = new File(pdfFile);      
            if(tofile.exists()) {      
                tofile.delete();      
            }      
            Dispatch.call(ppt, "SaveAs", pdfFile, PPT_FORMAT_PDF); 
            long end = System.currentTimeMillis();      
            System.out.println("转换完成..用时：" + (end - start) + "ms.");  
        } catch (Exception e) {      
            System.out.println("========Error:文档转换失败：" + e.getMessage());      
        } finally {  
            Dispatch.call(ppt, "Close");  
            System.out.println("关闭文档");  
            if (app != null)      
                app.invoke("Quit", new Variant[] {});      
            }  
          //如果没有这句话,winword.exe进程将不会关闭  
           ComThread.Release();
    }
	
	/**
	 * @Description:获取文件后缀
	 * @author xueyya
	 * @date:2018年3月16日 上午9:01:53
	 * @param fileName
	 * @return String
	 */
	private static String getFileSufix(String fileName) {
	    int splitIndex = fileName.lastIndexOf(".");
	    return fileName.substring(splitIndex + 1);
	}
	
}
