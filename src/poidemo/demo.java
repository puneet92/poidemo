package poidemo;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class demo {

	public static void main(String[] args) {
		
		 File file = new File("D:/Hadoop.pptx");
		  FileOutputStream out;
		  
		  try{
			  
			  FileInputStream inputstream = new FileInputStream(file);
		      XMLSlideShow ppt = new XMLSlideShow(inputstream);
		     
		      out = new FileOutputStream(file);
		      
		      //adding slides to the slodeshow
		      //adding slides to the slodeshow
		      XSLFSlide slide1 = ppt.createSlide();
		      XSLFSlide slide2 = ppt.createSlide();
		      ppt.write(out);
		      
		      System.out.println("Presentation edited successfully");
		      out.close();
		      
		  }catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
	      
	      
	  
		/*File file = new File("D:/examples1.pptx");
		FileOutputStream out;
		try {
			out = new FileOutputStream(file);
			ppt.write(out);
			System.out.println("document created sucessfully");
			out.close();
		} catch (IOException e) {
			System.out.println(e);	
			
		}*/
				
		
	}

}
