package poidemo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.sl.usermodel.Background;
import org.apache.poi.sl.usermodel.TextRun.FieldType;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;

public class poidemo3 {
	
	public static void main(String[]args){
		
		 File file = new File("D:/Hadoop.pptx");
		  FileOutputStream out;
		  try
		  {
			  FileInputStream inputstream = new FileInputStream(file);
		      XMLSlideShow ppt = new XMLSlideShow(inputstream);
		      /*for (XSLFSlideMaster master : ppt.getSlideMasters()){
		    	  for (XSLFSlideLayout layout : master.getSlideLayouts()){
		    	    if (layout1.getName().equals("Scorecard")) {
		    	        detailedscorecard=layout1;
		    	    }
		    		  System.out.println("Name: "+layout.getName()+" Type: "+layout.getType());
		    	  }
		    	 }*/
		      XSLFSlideMaster defaultMaster = ppt.getSlideMasters().get(0);
		     XSLFSlideLayout layout 
			  = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
		     //   ppt.getSlideMasters().get(0).getBackground().setFillColor(IndexedColors.YELLOW.getIndex());
		     ppt.getSlideMasters().get(0).getBackground().setFillColor(IndexedColors.GREY_50_PERCENT.getIndex());

		      //  ppt.getSlideMasters().get(0).getBackground().getFillPaint().setFillType(FieldType.Solid);
		       // ppt.getSlideMasters().get(0).getBackground().getFillPaint().getSolidFillColor().setColor(Color.GREEN);
				
		    //  out = new FileOutputStream(file);
		    //  ppt.write(out);
		      System.out.println("hi");
		      
		     // out.close();
		     
		     // out = new FileOutputStream(file);
		  }
		  catch(IOException e)
		  {
			  System.out.println("catch block");
		  }
	}

}
