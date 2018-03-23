package com.yuanzi.office2pdf;

import java.io.File;
import java.io.UnsupportedEncodingException;
import java.util.ResourceBundle;

import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;

/**
 * @Description:jodconverter + openOffice 
 * 				（一般格式实现效果还行，复杂格式容易有错位，跨平台）
 * 				必须安装openOffice
 * @author xueyya
 * @date:2018年3月16日 上午9:26:57
 */
public class JodConverter {
	
	/**
	 * @Description:根据文件类型转换为pdf
	 * @author xueyya
	 * @date:2018年3月16日 上午9:16:47
	 * @param inputFile
	 * @param pdfFile void
	 * @throws UnsupportedEncodingException 
	 */
	public static void convert2PDF(File inputFile, File pdfFile) {
		OfficeManager officeManager = null;
		try {
			long start = System.currentTimeMillis();
			DefaultOfficeManagerConfiguration config = new DefaultOfficeManagerConfiguration();  
			ResourceBundle resource = ResourceBundle.getBundle("application");
			String officeHome = resource.getString("office.home");
			officeHome = new String(officeHome.getBytes("ISO-8859-1"), "utf-8");
			config.setOfficeHome(officeHome);  
			officeManager = config.buildOfficeManager();  
			officeManager.start();  
			OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager); 
			System.out.println("转换文档到PDF..." + pdfFile.getPath());
			converter.convert(inputFile, pdfFile);
			long end = System.currentTimeMillis();
			System.out.println("转换完成..用时：" + (end - start) + "ms.");
		} catch (Exception e) {
			System.out.println(e.getMessage());
		} finally {
			if (officeManager != null) {
				officeManager.stop();
			}
		}
	}
	
}
