package org.geez.convert.docx;

/*
 * The non-maven way to build the jar file:
 *
 * javac -Xlint:deprecation -cp docx4j-6.0.1.jar:dependencies/commons-io-2.5.jar:../icu4j-63_1.jar:dependencies/slf4j-api-1.7.25.jar:slf4j-1.7.25 *.java
 * jar -cvf convert.jar org/geez/convert/docx/*.class org/geez/convert/tables/
 * java -cp convert.jar:docx4j-6.0.1.jar:dependencies/*:../icu4j-63_1.jar:slf4j-1.7.25/slf4j-nop-1.7.25.jar org.geez.convert.docx.ConvertDocx brana myFile-In.docx myFile-Out.docx
 *
 */

import org.docx4j.TraversalUtil;

import org.docx4j.XmlUtils;
import org.docx4j.finders.ClassFinder;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.FootnotesPart;
// import org.docx4j.openpackaging.parts.WordprocessingML.EndnotesPart;
import org.docx4j.openpackaging.parts.JaxbXmlPart;
import org.docx4j.wml.PPrBase.Spacing;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.PPr;
import org.docx4j.wml.STLineSpacingRule;
import org.docx4j.wml.CTSignedHpsMeasure;

//import org.docx4j.wml.R;
//import org.docx4j.wml.RPr;
//import org.docx4j.wml.RFonts;
//import org.docx4j.wml.Text;
import org.docx4j.wml.*;

//import java.io.BufferedReader;
//import java.io.File;
//import java.io.IOException;
//import java.io.InputStream;
//import java.io.InputStreamReader;
import java.util.List;
import java.util.Arrays;
import java.util.Scanner; 
import java.io.*;
import java.nio.file.*;
import java.nio.ByteBuffer;
import com.ibm.icu.text.*;
import java.math.*;
import org.apache.commons.lang3.StringEscapeUtils;
import java.lang.*;

public class ConvertDocx {
	protected Transliterator t = null;

	int Encodings = 17; //How many Encodings are there!!!!!  //2007  20080924
	int tcount = 0;
//	long [][] Table = new long [Encodings][];
	String [][] Table = new String [Encodings][];

	int tableSize[]= {0,30000,40000,135000,176000,1000,1000,40000,50000,90000,1000,40000,170000,680000,48000,40000};
	
	long[] start = new long[Encodings];
	long[] end = new long[Encodings];
	

	String  EncodingNames[] = {
		    "Unicode",      //ID--0
		    "TMW",          //ID--1
		    "TM",           //ID--2
		    "Fz",           //ID--3
		    "Hg",           //ID--4
		    "ACIP",         //ID--5
		    "Wylie",        //ID--6
		    "LTibetan",     //ID--7
		    "OldSambhota",  //ID--8
		    "NewSambhota",  //ID--9
		    "THDLWylie",    //ID--10
		    "LCWylie",      //ID--11
		    "TCRCBodYig",   //ID--12
		    "Bzd",          //ID--13 //2007
		    "Ty",           //ID--14
		    "NS",           //ID--15
		    "Jamyang"       //ID--16 //20080924
	};
	
	String EncodingFile[] = {
		    "none",      //ID--0
		    "TMW2Uni.tbl",          //ID--1
		    "TM2Uni.tbl",           //ID--2
		    "Fz2Uni.tbl",           //ID--3
		    "Hg2Uni.tbl",           //ID--4
		    "ACIP2Uni.tbl",         //ID--5
		    "Wylie2Uni.tbl",        //ID--6
		    "LTibetan2Uni.tbl",     //ID--7
		    "OldSambhota2Uni.tbl",  //ID--8
		    "NewSambhota2Uni.tbl",  //ID--9
		    "THDLWylie2Uni.tbl",    //ID--10
		    "LCWylie2Uni.tbl",      //ID--11
		    "TCRCBodYig2Uni.tbl",   //ID--12
		    "Bzd2Uni.tbl",          //ID--13 //2007
		    "Ty2Uni.tbl",           //ID--14
		    "NS2Uni.tbl",           //ID--15
		    "Jamyang2Uni.tbl"       //ID--16 //20080924
	};


	String[][] TibetanFontNames = {
		    //Unicode font
			{ "Microsoft Himalaya",
		    "SambhotaUnicode", //20181201
		    "CTRC-Uchen",  //20090325
		    "CTRC-Betsu",  //20090325
		    "CTRC-Drutsa",  //20090325
		    "CTRC-Tsumachu",  //20090325
		    "²ØÑÐÎÚ¼áÌå",  //20090325
		    "²ØÑÐ´ØÂêÇðÌå",  //20090325
		    "²ØÑÐ°Ø´ØÌå",  //20090325
		    "²ØÑÐÖé²ÁÌå",  //20090325
		    "²ØÑÐÎÚ½ðÌå"},  //20090325
		    //TMW font
			{"TibetanMachineWeb",
		    "TibetanMachineWeb1",
		    "TibetanMachineWeb2",
		    "TibetanMachineWeb3",
		    "TibetanMachineWeb4",
		    "TibetanMachineWeb5",
		    "TibetanMachineWeb6",
		    "TibetanMachineWeb7",
		    "TibetanMachineWeb8",
		    "TibetanMachineWeb9"},
		    //TM font
			{"TibetanMachine",
		    "TibetanMachineSkt1",
		    "TibetanMachineSkt2",
		    "TibetanMachineSkt3",
		    "TibetanMachineSkt4"},
			{},		    //Fz
		    {},//Hg

		    //ACIP
		    {"Arial",
		    "Times New Roman"},
		    //Wylie
		    {"Arial",
		    "Times New Roman"},
		    //LTibetan
		    {"LTibetan",
		    "LMantra"},
		    //OldSambhota
		    {"Sama",
		    "Samb",
		    "Samc",
		    "Esama",
		    "Esamb",
		    "Esamc"},
		    //NewSambhota
		    {"Dedris-a",
		    "Dedris-a1",
		    "Dedris-a2",
		    "Dedris-a3",
		    "Dedris-b",
		    "Dedris-b1",
		    "Dedris-b2",
		    "Dedris-b3",
		    "Dedris-c",
		    "Dedris-c1",
		    "Dedris-c2",
		    "Dedris-c3",
		    "Dedris-d",
		    "Dedris-d1",
		    "Dedris-d2",
		    "Dedris-d3",
		    "Dedris-e",
		    "Dedris-e1",
		    "Dedris-e2",
		    "Dedris-e3",
		    "Dedris-f",
		    "Dedris-f1",
		    "Dedris-f2",
		    "Dedris-f3",
		    "Dedris-g",
		    "Dedris-g1",
		    "Dedris-g2",
		    "Dedris-g3",
		    "Dedris-syma",
		    "Dedris-vowa",
		    "Ededris-a",
		    "Ededris-a1",
		    "Ededris-a2",
		    "Ededris-a3",
		    "Ededris-b",
		    "Ededris-b1",
		    "Ededris-b2",
		    "Ededris-b3",
		    "Ededris-c",
		    "Ededris-c1",
		    "Ededris-c2",
		    "Ededris-c3",
		    "Ededris-d",
		    "Ededris-d1",
		    "Ededris-d2",
		    "Ededris-d3",
		    "Ededris-e",
		    "Ededris-e1",
		    "Ededris-e2",
		    "Ededris-e3",
		    "Ededris-f",
		    "Ededris-f1",
		    "Ededris-f2",
		    "Ededris-f3",
		    "Ededris-g",
		    "Ededris-g1",
		    "Ededris-g2",
		    "Ededris-g3",
		    "Ededris-syma",
		    "Ededris-vowa"},
		    //THDLWylie
		    {"Arial",
		    "Times New Roman"},
		    //LCWylie
		    {"Arial",
		    "Times New Roman"},
		    //TCRCBodYig
		    {"TCRC Bod-Yig",
		    "TCRC Youtso",
		    "TCRC Youtsoweb"},
		    //Bzd
		    {"BZDBT", //2007
		    "BZDHT",
		    "BZDMT"},
		    //Ty
		    {"TIBETBT",
		    "TIBETHT",
		    "CHINATIBET"},//20090325
		    //NS
		    {"²ØÎÄÎá¼áÇíÌå"},
		    //Jamyang   20080924
		    {"DBu-can",
		    "Tibetisch dBu-can",
		    "Tibetisch dBu-can Overstrike"}
		};//TO BE EXTENDED
	String[][] AllTibetanFontNames = {
		    //Unicode font
			{"Microsoft Himalaya", //2007
		    "SambhotaUnicode", //20181201
		    "CTRC-Uchen",  //20090325
		    "CTRC-Betsu",  //20090325
		    "CTRC-Drutsa",  //20090325
		    "CTRC-Tsumachu",  //20090325
		    "²ØÑÐÎÚ¼áÌå",  //20090325
		    "²ØÑÐ´ØÂêÇðÌå",  //20090325
		    "²ØÑÐ°Ø´ØÌå",  //20090325
		    "²ØÑÐÖé²ÁÌå",  //20090325
		    "²ØÑÐÎÚ½ðÌå"},  //20090325
		    //TMW font
			{"TibetanMachineWeb",
		    "TibetanMachineWeb1",
		    "TibetanMachineWeb2",
		    "TibetanMachineWeb3",
		    "TibetanMachineWeb4",
		    "TibetanMachineWeb5",
		    "TibetanMachineWeb6",
		    "TibetanMachineWeb7",
		    "TibetanMachineWeb8",
		    "TibetanMachineWeb9"},
		    //TM font
			{"TibetanMachine",
		    "TibetanMachineSkt1",
		    "TibetanMachineSkt2",
		    "TibetanMachineSkt3",
		    "TibetanMachineSkt4"},
			{},//Fz
			{},//Hg
			{},//Chinese
			{},//"ËÎÌå", //20080229  //20090325
			//LTibetan
			{"LTibetan",
		    "LMantra"},
		    //OldSambhota
			{"Sama",
		    "Samb",
		    "Samc",
		    "Esama",
		    "Esamb",
		    "Esamc"},
		    //NewSambhota
			{"Dedris-a",
		    "Dedris-a1",
		    "Dedris-a2",
		    "Dedris-a3",
		    "Dedris-b",
		    "Dedris-b1",
		    "Dedris-b2",
		    "Dedris-b3",
		    "Dedris-c",
		    "Dedris-c1",
		    "Dedris-c2",
		    "Dedris-c3",
		    "Dedris-d",
		    "Dedris-d1",
		    "Dedris-d2",
		    "Dedris-d3",
		    "Dedris-e",
		    "Dedris-e1",
		    "Dedris-e2",
		    "Dedris-e3",
		    "Dedris-f",
		    "Dedris-f1",
		    "Dedris-f2",
		    "Dedris-f3",
		    "Dedris-g",
		    "Dedris-g1",
		    "Dedris-g2",
		    "Dedris-g3",
		    "Dedris-syma",
		    "Dedris-vowa",
		    
		    "Ededris-a",
		    "Ededris-a1",
		    "Ededris-a2",
		    "Ededris-a3",
		    "Ededris-b",
		    "Ededris-b1",
		    "Ededris-b2",
		    "Ededris-b3",
		    "Ededris-c",
		    "Ededris-c1",
		    "Ededris-c2",
		    "Ededris-c3",
		    "Ededris-d",
		    "Ededris-d1",
		    "Ededris-d2",
		    "Ededris-d3",
		    "Ededris-e",
		    "Ededris-e1",
		    "Ededris-e2",
		    "Ededris-e3",
		    "Ededris-f",
		    "Ededris-f1",
		    "Ededris-f2",
		    "Ededris-f3",
		    "Ededris-g",
		    "Ededris-g1",
		    "Ededris-g2",
		    "Ededris-g3",
		    "Ededris-syma",
		    "Ededris-vowa"},    

			{"TCRC Bod-Yig",
		    "TCRC Youtso",
		    "TCRC Youtsoweb"},
		    //Bzd
			{"BZDBT", //2007
		    "BZDHT",
		    "BZDMT"},
		    //Ty
			{"TIBETBT",
		    "TIBETHT",
		    "CHINATIBET"},//20090325
		    //NS
			{"²ØÎÄÎá¼áÇíÌå"},
		    //Jamyang   20080924
			{"DBu-can"},
			{"Tibetisch dBu-can",
		    "Tibetisch dBu-can Overstrike"}
		};//TO BE EXTENDED
	int TotalTibetanFontNumber = 107; // = 95 - 8 + 1; //# of all Tibetan fonts over all encodings  //2007           20080924 //20090325 //was 106 before 20181201
    //TO BE DETERMINED
//--------------------------------------------------------------------------
//Tibetan encodings. Index is the ID of the encoding

	//int BaseIndex[]={0, 10, 20, 0, 0, 25, 27, 29, 31, 37, 97, 99, 101, 104, 107, 110, 111};//Actual index of each  encoding in the array "TibetanFontNames"  //2007  20080924 //20090325
	int[] BaseIndex={0, 11, 21, 0, 0, 26, 28, 30, 32, 38, 98, 100, 102, 105, 108, 111, 112};//Actual index of each   encoding in the array "TibetanFontNames"  //2007  20080924 //20090325 //add SambhotaUnicode on 20181201


	//int EncodingFontNumber[] = {10, 10, 5, 0, 0, 2, 2, 2, 6, 60, 2, 2, 3, 3, 3, 1, 3}; //Number of Tibetan font in a certain font encoding //2007 20080924  //20090325
	int[] EncodingFontNumber = {11, 10, 5, 0, 0, 2, 2, 2, 6, 60, 2, 2, 3, 3, 3, 1, 3}; //Number of Tibetan font in a certain font encoding //2007   20080924  //20090325 // add SambhotaUnicode on 20181201

	public Boolean isTibetan(String fontname) {
		for(int m =0; m < Encodings; m++) {
//			System.out.print("font num="+EncodingFontNumber[m]+"\n");
			for (int n=0; n < EncodingFontNumber[m]; n++) {
//				System.out.print(TibetanFontNames[m][n]+"\n");
				if (TibetanFontNames[m][n].equals(fontname)) {
//					System.out.print("found font\n");
//					fonti=m;
//					fontFileID=n;
					return true;
//					break compare_font_loop;
				}
			}
		}
		return false;
	}
	public int whichTfont(String fontname) {
//		compare_font_loop:
		for(int m =0; m < Encodings; m++) {
//			System.out.print("font num="+EncodingFontNumber[m]+"\n");
			for (int n=0; n < EncodingFontNumber[m]; n++) {
//				System.out.print(TibetanFontNames[m][n]+"\n");
				if (TibetanFontNames[m][n].equals(fontname)) {
//					System.out.print("found font\n");
//					fonti=m;
//					fontFileID=n;
					return m;
//					break compare_font_loop;
				}
			}
		}
		return 0;
	}

	public String readRules( String fileName ) throws IOException {
		String line, segment, rules = "";

		ClassLoader classLoader = this.getClass().getClassLoader();
		InputStream in = classLoader.getResourceAsStream( "tables/" + fileName ); 
		BufferedReader ruleFile = new BufferedReader( new InputStreamReader(in, "UTF-8") );
		while ( (line = ruleFile.readLine()) != null) {
			if ( line.trim().equals("") || line.charAt(0) == '#' ) {
				continue;
			}
			segment = line.replaceFirst ( "^(.*?)#(.*)$", "$1" );
			rules += ( segment == null ) ? line : segment;
		}
		ruleFile.close();
		return rules;
		
	}

	
	public void readTables() throws IOException {

//		System.out.print("user.dir"+"\n");
//		String elements[];
//		BigInteger integer;
//		String line;
//		char ch;
//		int n;

		for(int i = 1; i < Encodings; i++) {
			String filename  = System.getProperty("user.dir") + "/" + EncodingFile[i];
//			System.out.println(System.getProperty("file.encoding"));
			//Scan long
//1			Scanner scanner = new Scanner(System.in); 
			if(i==5||i==6||i==10||i==11||i==13||i==14||i==15) {// for Bzd, Ty or NS, cannot convert now, need sample, load line 12380 in LoadMappingTableOthers2Unicode
			    System.out.println("Cannot convert Bzd, Ty or NS to unicode, if you provide sample file, I may add them.\n");	
			}
			else // For other Encodings such as Fz, Hg, LTibetan, ....
		    {
				// wrap a BufferedReader around FileReader
				BufferedReader bufferedReader = new BufferedReader(new FileReader(filename));
				
				// use the readLine method of the BufferedReader to read one line at a time.
				// the readLine method returns null when there is nothing else to read.
//				line = bufferedReader.readLine();
				Table[i] = bufferedReader.readLine().split(" ");
				
//				long [][] Table = new long [i][elements.length];
//				System.out.print("tableSize="+tableSize[i]+"\n");
				System.out.print(filename+"\n");

				// close the BufferedReader when we're done
				bufferedReader.close();
		    }
		}
	}
		
	
	public String convertText( String txt , String fnt) {
		
        int fonti=0;
        int fontID=0;
        int l=-1;
        int val;
//        int fontFileID;
	    char [] charArr = new char[4];
	    char ch;
	    String out="";
	    String temp="";
	    String unicode;
	    String text;
	    String font;
	    String result="";

	    
	    String unicodes;

	    text=txt;
	    font=fnt;
//	    ByteBuffer uni = ByteBuffer.allocate(8);

	    System.out.println("text len="+text.length()+"\n");
		for(int i = 0; i < text.length(); i++) {
//			int x =  ( 0x00ff & (int)text.charAt(i) );
//			int x =  (int)text.charAt(i) ;
			int x =  text.charAt(i) ;
			System.out.print("x="+x+"\n");
    
			String hex=Integer.toHexString(x);
			int value = Integer.parseInt(hex, 16);
			int diff=value-0x21;
			System.out.print("diff="+diff+"\n");
//			if(value < 0x21 && value > 0xff) continue;
			out="";
	     if(x >= 0x21 && x <= 0xff) {
//		  if(x >= 33 || x <= 255) {
	    	 System.out.print("still entered\n");
			StringBuilder sb = new StringBuilder();
			if (value != 0x20) { // value =32, this is an spaceless space
//				out="";
				tcount++;
				System.out.print("tcount="+tcount+"\n");
			}else {
				continue;
			}
			System.out.print("hex="+hex+"\n");
			System.out.print("value1="+value+"\n");
//			System.out.print("value="+value+"\n");
//			System.out.print((int)text.charAt(i)+"\n");
//			System.out.print(hex+"\n");
//			System.out.print("font="+font+"\n");
			compare_font_loop:
			for(int m =0; m < Encodings; m++) {
//				System.out.print("font num="+EncodingFontNumber[m]+"\n");
				for (int n=0; n < EncodingFontNumber[m]; n++) {
//					System.out.print(TibetanFontNames[m][n]+"\n");
					if (TibetanFontNames[m][n].equals(font)) {
//						System.out.print("found font\n");
						fonti=m;
						fontID=n;
						break compare_font_loop;
					}
				}
			}
/*			if(fonti==0||fonti==5||fonti==6||fonti==10||fonti==11||fonti==13||fonti==14||fonti==15) {// for Bzd, Ty or NS, cannot convert now, need sample, load line 12380 in LoadMappingTableOthers2Unicode
			    System.out.println("Cannot convert Bzd, Ty or NS to unicode, if you provide sample file, I may add them.\n");	
			}
*/	
			sb.append(  (char)x );
			System.out.print("sb="+sb + "\n");
//			System.out.print((int)text.charAt(i)+"\n");
			System.out.print("Chosen\n");
//		    "Unicode", ID--0;  "TMW", ID--1; "TM", ID--2; "Fz",ID--3;  "Hg",ID--4;  "ACIP", ID--5
//"Wylie", ID--6;  "LTibetan", ID--7;  "OldSambhota", ID--8;  "NewSambhota", ID--9;  
//"THDLWylie", ID--10;  "LCWylie",  ID--11;   "TCRCBodYig",  ID--12; "Bzd", ID--13;
//"Ty", ID--14; "NS", ID--15; "Jamyang" ID--16 			
		    System.out.print("fonti="+fonti+"\n");

			if(fonti==9 && fontID > 29) { //If NewSambhota
				fontID = fontID - 30;
			}else if(fonti==8 && fontID > 2) { // if OldSambhota
				fontID = fontID - 3;
				 
			}
		    if(fonti==1 || fonti==9) {// If TMW or NewSambhota
		       l = fontID * 94 + (value - 0x21);
				System.out.print("value2="+value+"\n");
			    System.out.print("l="+l+"\n");
		    }
		    else if(fonti == 16) { // If Jamyang
		    	l = fontID * 223 + (value - 0x21); 
		    }
		    else {
		    	l = fontID * 222 + (value - 0x21);
		    }			 
			System.out.print("fontiD="+fontID+"\n");
//			if(l==-1) {
//				out=" ";
//			}
			if(fonti==9) {
//		       int thepoint=l*30;
//			       long thepoint=start[fonti]+l*30;
//			       int n=0;
//		    	String[] temp=new String[5];
//		    	String temp;
			       System.out.println("Len="+Table[fonti].length+"\n");
			       unicodes=Table[fonti][l];
			       System.out.println("len="+unicodes.length()+"\n");
			       System.out.print("bigInt="+unicodes+"\n");
			       int n=0;
			       out="";
			       temp="";
			       unicode="";
			       for (int k=0;k<unicodes.length();k++)
			        {
			            ch =  unicodes.charAt(k);
		            	temp+=ch;
			            n++;
			            if(n==4)
			            {
			            	unicode="\\u"+temp;
			            	val = Integer.parseInt(temp);	
			            	hex = Integer.toHexString(val);
			            	hex="\\u0"+hex;
			            	ch = (char) Integer.parseInt( hex.substring(2), 16 );
//			            	System.out.print("hex="+hex+"\n");
			            	System.out.print("ch="+ch+"\n");
//			        		Integer code = Integer.parseInt(unicode.substring(2), 16); // the integer 65 in base 10
//			        		ch = Character.toChars(code)[0]; // the letter 'A'
			            	out+=ch;
			                n=0;
			                temp="";
			                unicode="";
			            }
			        }			       
	                System.out.println("out="+out+"\n");
			}
			else {
//						out=text;
				System.out.println("This shouldn't happen\n");
				System.exit(1);

			}
		
		  }
		  else if(x == 0x20){
			  if(text.length()==1) {
			  hex="\\u0020"; // this is a space
          	ch = (char) Integer.parseInt( hex.substring(2), 16 );
          	out += ch;
			  }else  {
//				  ch='';
//		          	out += ch;
			  }
		  }
  		  else if(x == 0x09){
            hex="\\u0009"; // this is a space
            ch = (char) Integer.parseInt( hex.substring(2), 16 );
            out += ch;
		  }else {
            out+=text.charAt(i);
		  }
            System.out.println("out1="+out+"\n");
            result=result+out;
		}
//		String step1 = t.transliterate( sb.toString() );
//		String step2 = (step1 == null ) ? null : step1.replaceAll( "፡፡", "።"); // this usually won't work since each hulet neteb is surrounded by separate markup.
//		return step2;
		//String s = StringEscapeUtils.unescapeJava(out); // s contains the euro symbol followed by newline
		//String charInUnicode = "\\u0041"; // ascii code 65, the letter 'A'
//		ch = (char) code;
//		String s=Character.toString(ch);
//		System.out.println(s);
//	    out="done";
        System.out.println("out1.5="+result+"\n");
	    return result;
//		return ch;
	}



	public void processObjects(
		final JaxbXmlPart<?> part,
		final Transliterator translit1,
		final Transliterator translit2,
		final String fontName1,
		final String fontName2) throws Docx4JException
	{
		Boolean tibetan=false;	
//		String text;
//		String font;
		int fonti=0;
		BigInteger bigInt;
		BigInteger eight = new BigInteger("8");
//		Object p2;
//		CTSignedHpsMeasure posi== docx.getFactory().createCTSignedHpsMeasure();;					

			    
		ClassFinder pfinder = new ClassFinder( P.class );
		new TraversalUtil(part.getContents(), pfinder);
		for (Object o : pfinder.results) {
			Object o2 = XmlUtils.unwrap(o);
			if (o2 instanceof org.docx4j.wml.P) {
				P p = (org.docx4j.wml.P)o2;
			
//				P  paragraph = ((P)XmlUtils.unwrap(o) );
			    
//				PPr paragraphProperties = factory.createPPr();
//		           wmlP = documentPart.createParagraphOfText(null);				
				// this is ok, provided the results of the Callback
				// won't be marshalled			
				ClassFinder rfinder = new ClassFinder( R.class );
				new TraversalUtil(p.getContent(), rfinder);
				for (Object o3 : rfinder.results) {


					if (o3 instanceof org.docx4j.wml.R) {
						R r = (org.docx4j.wml.R)o3;

						RPr rpr = r.getRPr();

						//if(rpr.getPosition()!=null) {
//						CTSignedHpsMeasure posi=rpr.getPosition();
//						posi.setVal(BigInteger.ZERO);
//					}
					 
						if (rpr == null ) continue;
						RFonts rfonts = rpr.getRFonts(); 
//						System.out.print("EastAsian="+rpr.getEastAsianLayout().toString()+"\n");
//						System.out.print("Style="+rpr.getRStyle().getVal()+"\n");

						HpsMeasure size = rpr.getSz();
						if( rfonts == null ) {
							tibetan = false;
							System.out.println("find null font\n");
						}
						else{
							fonti=whichTfont(rfonts.getAscii());
							System.out.print("font="+rfonts.getAscii()+"\n");
				
//							if(fonti!=0&&fonti!=5&&fonti!=6&&fonti!=10&&fonti!=11&&fonti!=13&&fonti!=14&&fonti!=15) {// for Bzd, Ty or NS, cannot convert now, need sample, load line 12380 in LoadMappingTableOthers2Unicode
							if(fonti==9) {// can only convert NewSambhota to Unicode right now, load line 12380 in LoadMappingTableOthers2Unicode
								tibetan=true;							
							}
						/*
						else if() {
							rfonts.setAscii( "SambhotaUnicode" );
							rfonts.setHAnsi( "SambhotaUnicode" );
							rfonts.setCs( "SambhotaUnicode" );
							rfonts.setEastAsia( "SambhotaUnicode" );
							tibetan=true;
						}*/
						/*						else if( fontName1.equals( rfonts.getAscii() ) ) {
							rfonts.setAscii( "Abyssinica SIL" );
							rfonts.setHAnsi( "Abyssinica SIL" );
							rfonts.setCs( "Abyssinica SIL" );
							rfonts.setEastAsia( "Abyssinica SIL" );
							t = translit1;
						}
						else if( fontName2.equals( rfonts.getAscii() ) ) {
							/*rfonts.setAscii( "Abyssinica SIL" );
							rfonts.setHAnsi( "Abyssinica SIL" );
							rfonts.setCs( "Abyssinica SIL" );
							rfonts.setEastAsia( "Abyssinica SIL" );
							t = translit2;
						}*/
//						else {
//							t = null;
//							tibetan=false;
//						}
						}
						t=translit1; //20181215 add for debug
					
						List<Object> objects = r.getContent();
						for ( Object x : objects ) {
							Object x2 = XmlUtils.unwrap(x);
		                        
							if ( x2 instanceof org.docx4j.wml.Text ) {
//								if ( tibetan == true) {
								if ( fonti == 9) {
								
									Text txt = (org.docx4j.wml.Text)x2;
//									text=txt.getValue();
//									font=rfonts.getAscii();
									System.out.print("txt="+txt.getValue()+"\n");
									
									String out = convertText( txt.getValue() , rfonts.getAscii());
//									String out = convertText( text , font);
									System.out.print("out2="+out+"\n");
									txt.setValue( out );
									if ( " ".equals( out ) ) {	
										txt.setSpace( "preserve" );
									}
									if(size != null) {//NewSambhota font size decrease 4
//										size.setVal
										bigInt=size.getVal();
										System.out.print("fontsize"+bigInt+"\n");
										size.setVal(bigInt.subtract(eight));
									}
										PPr pPr = p.getPPr();

										if (pPr==null) {
											System.out.print("no spacing properties\n");
							                pPr = Context.getWmlObjectFactory().createPPr();
//											paragraphProperties.getSpacing().setBeforeLines(BigInteger.valueOf(10));
										}
										Spacing spacing=pPr.getSpacing();
										if(spacing==null) {
									          spacing = Context.getWmlObjectFactory().createPPrBaseSpacing();
//											  spacing.setLine(BigInteger.valueOf(500));
											  spacing.setBefore(BigInteger.valueOf(100));
									          System.out.print("no spacing\n");
										}
											//								    paragraphProperties.getSpacing().setLine(BigInteger.valueOf(500));
//									    Spacing sp = factory.createPPrBaseSpacing();
									    
//									    sp.setAfter(BigInteger.valueOf(200));
//										    sp.setLine(BigInteger.valueOf(5));
//										    sp.setLine(BigInteger.ZERO);
//									    sp.setLineRule(STLineSpacingRule.AUTO);
									    pPr.setSpacing(spacing);
									    
										if(rpr.getPosition()==null) {
											CTSignedHpsMeasure posi=Context.getWmlObjectFactory().createCTSignedHpsMeasure();
											posi.setVal(BigInteger.valueOf(-6));
											rpr.setPosition(posi);

										}
									   
//									    Style style.setPPr(paragraphProperties);
									
//										ObjectFactory factory = Context.getWmlObjectFactory();
//								    	Spacing spacing = Context.getWmlObjectFactory().createPPrBaseSpacing();
//								    	documentDefaultPPr.setSpacing(spacing);
//								    	spacing.setBefore(BigInteger.valueOf(300));
//								    	spacing.setAfter(BigInteger.valueOf(300));
//								    	spacing.setLine(BigInteger.valueOf(240));
//								    	System.out.println("spacing="+spacing+"\n");
										rfonts.setAscii( "SambhotaUnicode" );
										rfonts.setHAnsi( "SambhotaUnicode" );
										rfonts.setCs( "SambhotaUnicode" );
										rfonts.setEastAsia( "SambhotaUnicode" );
										System.out.print("convert one\n");
								    
									}
//									ObjectFactory factory = (org.docx4j.wml.ObjectFactory) factory;
									
//								}
									tibetan=false;
									fonti=0;
								}
							
							else {
							// System.err.println( "Found: " + x2.getClass() );
							}
						}
					} else {
//						System.err.println( XmlUtils.marshaltoString(o, true, true) );
					}
				}
			  }
			}

		}



	public void process(
		final String table1RulesFile,
		final String table2RulesFile,
		final String fontName1,
		final String fontName2,
		final File inputFile,
		final File outputFile)
	{

		try {
			// specify the transliteration file in the first argument.
			// read the input, transliterate, and write to output
			String table1Text = readRules( table1RulesFile  );
			String table2Text = readRules( table2RulesFile );
			readTables();

			final Transliterator translit1 = Transliterator.createFromRules( "Ethiopic-ExtendedLatin", table1Text.replace( '\ufeff', ' ' ), Transliterator.REVERSE );
			final Transliterator translit2 = Transliterator.createFromRules( "Ethiopic-ExtendedLatin", table2Text.replace( '\ufeff', ' ' ), Transliterator.REVERSE );


			WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load( inputFile );		
			MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
            processObjects( documentPart, translit1, translit2, fontName1, fontName2 );
            
            if( documentPart.hasFootnotesPart() ) {
            	FootnotesPart footnotesPart = documentPart.getFootnotesPart();
            	processObjects( footnotesPart, translit1, translit2, fontName1, fontName2 );	
            }

   
			// Save it zipped
			wordMLPackage.save( outputFile );

		} catch ( Exception ex ) {
			System.err.println( ex );
		}
   

	}
	

	public static void main( String[] args ) {
		if( args.length != 3 ) {
			System.err.println( "Exactly 3 arguements are expected: <system> <input file> <output file>" );
			System.exit(0);
		}

		String system = args[0];
		String inputFilepath  = System.getProperty("user.dir") + "/" + args[1];
		String outputFilepath = System.getProperty("user.dir") + "/" + args[2];
		File inputFile = new File ( inputFilepath );
		File outputFile = new File ( outputFilepath );


		if( "brana".equals( system ) ) {
			ConvertDocx converter = new ConvertDocx();
			converter.process( "BranaITable.txt", "BranaIITable.txt", "Brana I", "Brana II", inputFile, outputFile );
		}
		else if( "geeznewab".equals( system ) ) {
			ConvertDocxFeedelGeezNewAB converter = new ConvertDocxFeedelGeezNewAB();
			converter.process( "GeezNewATable.txt", "GeezNewBTable.txt", "GeezNewA", "GeezNewB",  inputFile, outputFile );
		}
		else {
			System.err.println( "Unrecognized input system: " + system );
		}
	}
}
