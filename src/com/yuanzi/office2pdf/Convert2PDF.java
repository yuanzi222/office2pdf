package com.yuanzi.office2pdf;

import java.io.File;

/**
 * @Description:office转换为pdf
 * @author xueyya
 * @date:2018年3月16日 上午9:28:48
 */
public class Convert2PDF {
	
	public static void main(String[] args) {
		String inputFile = "e:\\1.docx";
//		String inputFile = "e:\\1.xlsx";
//		String inputFile = "e:\\1.pptx";
		String pdfFile = getToFileName(inputFile);
		String suffix = getFileSufix(inputFile);
	    File file = new File(inputFile);
	    if (!file.exists()) {
	        System.out.println("文件不存在！");
	    }
	    if (suffix.equals("pdf")) {
	        System.out.println("PDF not need to convert!");
	    }
	    //Jacob方式
//	    Jacob.convert2PDF(inputFile, pdfFile);
	    //JodConverter方式
	    convert2PDFForJodConverter(inputFile);
	}
	
	/**
	 * @Description:通过jodconverter + openOffice转换
	 * @author xueyya
	 * @date:2018年3月21日 下午2:27:36
	 * @param inputFileName void
	 */
	private static void convert2PDFForJodConverter(String inputFileName) {
		File inputFile = new File(inputFileName);
		File pdfFile = new File(getToFileName(inputFileName));
		JodConverter.convert2PDF(inputFile, pdfFile);
	}
	
	/**
	 * @Description:获取转换后的地址
	 * @author xueyya
	 * @date:2018年3月15日 下午4:07:30
	 * @param inputFile
	 * @return String
	 */
	private static String getToFileName(String inputFile) {
		int lastIndexOf = inputFile.lastIndexOf(".");
		String pdfFile = inputFile.substring(0, lastIndexOf) + ".pdf";
		return pdfFile;
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
