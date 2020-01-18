package poidemo;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;

public class poidemo2 {
	
	public static void main(String[]args)
	{
		
		
		XMLSlideShow ppt = new XMLSlideShow();
		//ppt.createSlide();
		XSLFSlideMaster defaultMaster = ppt.getSlideMasters().get(0);
		XSLFSlideLayout layout 
		  = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
		XSLFSlide slide = ppt.createSlide(layout);
		
		FileOutputStream out;
		try {
			out = new FileOutputStream("D:/powerpoint.pptx");
			ppt.write(out);
			out.close();
			System.out.println("hii");
		} catch ( IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
	

}
