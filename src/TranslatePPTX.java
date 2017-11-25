/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
package org.apache.poi.xslf.extractor;

import java.io.IOException;
import java.awt.Dimension;
import java.awt.Color;
import java.awt.Paint;
import java.awt.geom.Rectangle2D;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.BufferedReader;
import java.io.InputStream;
import java.io.PrintStream;

import java.util.*;

import java.lang.Enum;

import org.apache.poi.POIXMLTextExtractor;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.sl.draw.DrawPaint;
import org.apache.poi.sl.draw.DrawTextParagraph;
import org.apache.poi.sl.usermodel.PaintStyle;
import org.apache.poi.sl.usermodel.PaintStyle.SolidPaint;
import org.apache.poi.sl.usermodel.ColorStyle;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFCommentAuthors;
import org.apache.poi.xslf.usermodel.XSLFComments;
import org.apache.poi.xslf.usermodel.XSLFHyperlink;
import org.apache.poi.xslf.usermodel.XSLFNotes;
import org.apache.poi.xslf.usermodel.XSLFRelation;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFShapeContainer;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFSlideShow;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.presentationml.x2006.main.CTComment;
import org.openxmlformats.schemas.presentationml.x2006.main.CTCommentAuthor;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextCharacterProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTFontCollection;
import org.openxmlformats.schemas.drawingml.x2006.main.CTFontScheme;
import org.openxmlformats.schemas.drawingml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTSchemeColor;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeStyle;
import org.openxmlformats.schemas.drawingml.x2006.main.CTSolidColorFillProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextCharacterProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextField;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextFont;

public class TranslatePPTX extends POIXMLTextExtractor {
   public static final XSLFRelation[] SUPPORTED_TYPES = new XSLFRelation[] {
      XSLFRelation.MAIN, XSLFRelation.MACRO, XSLFRelation.MACRO_TEMPLATE,
      XSLFRelation.PRESENTATIONML, XSLFRelation.PRESENTATIONML_TEMPLATE,
      XSLFRelation.PRESENTATION_MACRO
   };

        public enum ModeValue { EXTRACT, TRANSLATE }
        public static ModeValue Mode = ModeValue.EXTRACT;
   
        public static int nText=0;
        public static final String TEXT_KEYWORD = "TEXT_STRING";
        public static final String TEXT_SEPARATOR = "==================================================";
        public static final Map <Dimension, String> Translations = new HashMap<Dimension, String>();
        public static final Map <Dimension, String> Formats = new HashMap<Dimension, String>();
        public static boolean Verbose=false;
        public static boolean Autosize=false;
        public static boolean WideOnly=false;
        public static boolean WriteFormats=false;
        public static boolean OmitRuns=false;
        public static FileWriter LogFile=null;
        public static FileWriter TextFile=null;
        public static int TotalTextShapes=0;
        public static int TotalTextRuns=0;
        public static int TotalTableEntries=0;
        public static int TextShapesReplaced=0;
        public static int TextRunsReplaced=0;

	private XMLSlideShow slideshow;
	private boolean slidesByDefault = true;
	private boolean notesByDefault;
        private boolean masterByDefault;
	
	public TranslatePPTX(XMLSlideShow slideshow) {
		super(slideshow);
		this.slideshow = slideshow;
	}
	public TranslatePPTX(XSLFSlideShow slideshow) throws XmlException, IOException {
		this(new XMLSlideShow(slideshow.getPackage()));
	}
	public TranslatePPTX(OPCPackage container) throws XmlException, OpenXML4JException, IOException {
		this(new XSLFSlideShow(container));
	}

	public static void Usage()
         { System.err.println("\nusage: TranslatePPTX Original.pptx [options]\n");
	   System.err.println(" options: \n");
	   System.err.println("  --Translations Translations.txt");
	   System.err.println("  --OutFile      OutputFile.pptx");
	   System.err.println("  --WideOnly");
	   System.err.println("  --OmitRuns");
	   System.err.println("  --Autosize");
	   System.err.println("  --WriteFormats");
	   System.err.println("  --Verbose");
	   System.err.println("  --WriteLog");
	   System.err.println("\n");
           System.exit(1);
         }

	public static void ErrExit(String msg)
         { 
           System.err.println("error: " + msg + " (aborting)");
           System.exit(1);
         }

        public static boolean HasWideCharacters(String s)
         { return !(s.chars().allMatch(c -> c < 128)); }

	public static void ReadTranslations(String FileName) throws Exception
         {
           try { FileReader fr     = new FileReader(FileName);
                 BufferedReader br = new BufferedReader(fr);

                 String Line;
                 int LineNum=0;
                 boolean OverlapsExist=false;
                 while ( (Line = br.readLine()) != null)
                  { 
                    LineNum++;

                    // 20171123 allow early termination
                    if (Line.equals("END"))
{ System.out.println("Breaking early on line " + LineNum);
                     break;
}

                    // skip if the line is not of the form TEXT_STRING NT NR
                    String Tokens[] = Line.split(" ");
                    if ( !(Tokens[0].equals(TEXT_KEYWORD) ) )
                     continue;
                    if ( Tokens.length < 3 )
                     ErrExit(FileName + ":" + LineNum + ":" + "too few tokens on line");
                    Dimension Key = new Dimension(0,0);
                    try
                     { 
                       Key.width=Integer.parseInt(Tokens[1]);
                       Key.height=Integer.parseInt(Tokens[2]);
                     }
                    catch(NumberFormatException e)
                     { 
                       ErrExit(FileName + ":" + LineNum + ":" + "syntax error");
                     }

                    // sanity check: no key exists in table twice
                    if (Translations.containsKey(Key))
                     ErrExit(FileName + ":" + LineNum + ":" + "Key (" + Key.width + "," + Key.height + ") exists twice");

                    // sanity check: for a given NT, *either* (NT,0) *or* (NT,i>0) may exist,
                    // but not both (each TextShape is either translated all-at-once or by individual TextRuns)
                    if (Key.height>=1)
                     { Dimension KeyPrime=new Dimension(Key);
                       KeyPrime.height=0;
                       if (Translations.containsKey(KeyPrime))
                        { OverlapsExist=true;
                          System.err.println(FileName + ":" + LineNum + ":" + "Key(" + Key.width + "," + Key.height + ") conflicts with existing (" + Key.width + ",0)");
                        };
                     };

                    // record format strings if any 
                    if (Tokens.length > 3)
                     { String Format = new String(Tokens[3]);
                       for(int nt=4; nt<Tokens.length; nt++)
                        Format = Format + " " + Tokens[nt];
                       Formats.put(Key, Format);
                     };

                    // next line must be text separator
                    Line=br.readLine();
                    if ( Line==null || !Line.equals(TEXT_SEPARATOR) )
                     ErrExit(FileName + ":" + LineNum + ":" + "expected separator string (=================), got " + Line);
                    LineNum++;

                    String NewText="";
                    // read lines up to next separator, then add (Key,NewText) pair to dictionary
                    for(;;)
                     {
                       Line = br.readLine();
                       LineNum++;

                       if (Line == null)
                        ErrExit(FileName + ":" + "unexpected end of file");
                       if (Line.startsWith(TEXT_KEYWORD,0))
                        ErrExit(FileName + ":" + LineNum + ": unterminated translation string");
                       if (Line.equals(TEXT_SEPARATOR))
                        break;

                       NewText = NewText + Line + "\n";
                     } 
                    Translations.put(Key, NewText);

                  }; // while ( (Line = br.readLine()) != null)

               if (LogFile!=null)
                { PrintLn(LogFile,"\n\n**Read " + Translations.size() + " replacement text strings: \n\n");
                  for (Map.Entry<Dimension, String> e : Translations.entrySet())
                   PrintLn(LogFile,"Text (" + e.getKey().width + "," + e.getKey().height + "): " + e.getValue());
                }
               if (OverlapsExist)
                ErrExit(FileName + "overlaps exist ");

               }
           catch(IOException e)
               { 
                  ErrExit("could not open file " + FileName);
               }
         };

	public static void Print(FileWriter f, String s)
	 {  
           if (f==null) return;
           try 
            { 
               f.write(s);
               f.flush();
            }
           catch(Exception e) 
            {
            }
         }

	public static void PrintLn(FileWriter f, String s)
	 { Print(f,s + "\n"); }

	public static void main(String[] args) throws Exception 
         {
           /***************************************************************/
           /* try to open PPTX file                                       */
           /***************************************************************/
           if(args.length < 1) Usage();
           String InputPPTXFile = args[0];
           String DirComponents[] = InputPPTXFile.split("/");
           int ndc=DirComponents.length;
           String FileBase = DirComponents[ndc-1].split(".pptx")[0];
           try 
               { 
                 FileInputStream InStream=new FileInputStream(InputPPTXFile);
                 InStream.close();
               }
           catch(Exception e)
               { 
                 ErrExit("could not open file " + InputPPTXFile);
               }
           XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(InputPPTXFile));

           /***************************************************************/
           /** parse command-line arguments *******************************/
           /***************************************************************/
           boolean WriteLog=false;
           String TextFileName=null, OutputPPTXFileName=null;
           for(int narg=1; narg<args.length; narg++)
            { 
              if ( args[narg].equalsIgnoreCase("--Verbose") )
               Verbose=true;
              else if ( args[narg].equalsIgnoreCase("--WriteLog") )
               WriteLog=true;
              else if ( args[narg].equalsIgnoreCase("--WideOnly") )
               WideOnly=true;
              else if ( args[narg].equalsIgnoreCase("--WriteFormats") )
               WriteFormats=true;
              else if ( args[narg].equalsIgnoreCase("--OmitRuns") )
               OmitRuns=true;
              else if ( args[narg].equalsIgnoreCase("--Translations") )
               { Mode=ModeValue.TRANSLATE;
                 ReadTranslations(args[narg+1]);
                 narg+=1;
               } 
              else if ( args[narg].equalsIgnoreCase("--PPTXOutput") )
               { 
                 OutputPPTXFileName = new String(args[narg+1]);
                 narg+=1;
               }
              else if ( args[narg].equalsIgnoreCase("--TextOutput") )
               { 
                 TextFileName = new String(args[narg+1]);
                 narg+=1;
               }
              else 
               ErrExit("unknown option " + args[narg]);
            };
          
           /***************************************************************/
           /* process PPTX file *******************************************/
           /***************************************************************/
           if (WriteLog)
            LogFile = new FileWriter(FileBase + ".log");

           if (Mode==ModeValue.EXTRACT)
            { if (TextFileName==null)
               TextFileName = new String(FileBase + ".text");
              TextFile = new FileWriter(TextFileName);
            };

           POIXMLTextExtractor extractor = new TranslatePPTX(ppt);
           extractor.getText();
	   extractor.close();

           if (Mode==ModeValue.EXTRACT)
            { TextFile.close();
              System.out.println("Wrote " + TotalTextShapes + " text strings to " + FileBase + ".text");
              System.out.println(" (" + TotalTextRuns + " text runs, " + TotalTableEntries + " table entries)");
            };

           /***************************************************************/
           /* Write translated PPTX file if we're in translate mode *******/
           /***************************************************************/
           if (Mode==ModeValue.TRANSLATE)
            { 
              System.out.println("Replaced " + TextShapesReplaced + " text shapes, " + TextRunsReplaced + " text runs");

              if (OutputPPTXFileName == null)
               OutputPPTXFileName = new String(FileBase + "_Translated.pptx");
              try (FileOutputStream PPTXFile = new FileOutputStream(OutputPPTXFileName) )
               {
                 ppt.write(PPTXFile);
                 System.out.println("Wrote translated document to " + OutputPPTXFileName + ".");
               }
              catch(Exception e)
               {
                 ErrExit("could not write output file " + OutputPPTXFileName);
               }
            }

	   ppt.close();
           System.out.println("Thank you for your support.");
	}

	/**
	 * Should a call to getText() return slide text?
	 * Default is yes
	 */
	public void setSlidesByDefault(boolean slidesByDefault) {
		this.slidesByDefault = slidesByDefault;
	}
	/**
	 * Should a call to getText() return notes text?
	 * Default is no
	 */
	public void setNotesByDefault(boolean notesByDefault) {
		this.notesByDefault = notesByDefault;
	}
	
   /**
    * Should a call to getText() return text from master? Default is no
    */
   public void setMasterByDefault(boolean masterByDefault) {
       this.masterByDefault = masterByDefault;
   }
	
	/**
	 * Gets the slide text, but not the notes text
	 */
	@Override
    public String getText() {
		return getText(slidesByDefault, notesByDefault);
	}
	
   /**
    * Gets the requested text from the file
    * @param slideText Should we retrieve text from slides?
    * @param notesText Should we retrieve text from notes?
    */
   public String getText(boolean slideText, boolean notesText) {
      return getText(slideText, notesText, masterByDefault);
   }
   
   /**
    * Gets the requested text from the file
    * 
    * @param slideText Should we retrieve text from slides?
    * @param notesText Should we retrieve text from notes?
    * @param masterText Should we retrieve text from master slides?
    * 
    * @return the extracted text
    */
   public String getText(boolean slideText, boolean notesText, boolean masterText) {

      int nSlide=0;
      for (XSLFSlide slide : slideshow.getSlides()) {

          nSlide=nSlide+1;
          
          PrintLn(TextFile,"\n");
          PrintLn(TextFile,"--------------------------------------------------");
          PrintLn(TextFile,"## Slide " + nSlide + ": " + slide.getTitle());
          PrintLn(TextFile,"--------------------------------------------------");

          if (Verbose)
           System.out.print(" Slide " + nSlide + ": ");

          StringBuilder SlideText = new StringBuilder();
          int TTS = TotalTextShapes;
          int TSR = TextShapesReplaced;
          int TRR = TextRunsReplaced;
          SlideText.append(getText(slide, slideText, notesText, masterText));
          TTS = TotalTextShapes - TTS;
          TSR = TextShapesReplaced - TSR;
          TRR = TextRunsReplaced - TRR;

          if (Verbose)
           { 
             if (Mode==ModeValue.EXTRACT )
               System.out.println(" extracted " + TTS + " text shapes ");
             else // Mode==ModeValue.TRANSLATE
               System.out.println(" replaced " + TSR + " text shapes, " + TRR + " text runs");
           }

          if (Mode==ModeValue.TRANSLATE && Translations.isEmpty())
           break;
          //System.out.println(SlideText.toString());
      }

      return null; //text.toString();
   }

   /**
    * Gets the requested text from the slide
    * 
    * @param slide the slide to retrieve the text from
    * @param slideText Should we retrieve text from slides?
    * @param notesText Should we retrieve text from notes?
    * @param masterText Should we retrieve text from master slides?
    * 
    * @return the extracted text
    */
   public static String getText(XSLFSlide slide, boolean slideText, boolean notesText, boolean masterText) {
       StringBuilder text = new StringBuilder();

       XSLFCommentAuthors commentAuthors = slide.getSlideShow().getCommentAuthors();

       XSLFNotes notes = slide.getNotes();
       XSLFComments comments = slide.getComments();
       XSLFSlideLayout layout = slide.getSlideLayout();
       XSLFSlideMaster master = layout.getSlideMaster();

       // TODO Do the slide's name
       // (Stored in docProps/app.xml)

       // Do the slide's text if requested
       if (slideText) {
          extractText(slide, false, text);
          
          // If requested, get text from the master and it's layout 
          if(masterText) {
             assert (layout != null);
             extractText(layout, true, text);
             assert (master != null);
             extractText(master, true, text);
          }

          // If the slide has comments, do those too
          if (comments != null) {
             for (CTComment comment : comments.getCTCommentsList().getCmArray()) {
                // Do the author if we can
                if (commentAuthors != null) {
                   CTCommentAuthor author = commentAuthors.getAuthorById(comment.getAuthorId());
                   if(author != null) {
                      text.append(author.getName() + ": ");
                   }
                }
                
                // Then the comment text, with a new line afterwards
                text.append(comment.getText());
                text.append("\n");
             }
          }
       }

       // Do the notes if requested
       if (notesText && notes != null) {
          extractText(notes, false, text);
       }
       
       return text.toString();
   }

    private static void Resize(XSLFTextShape ts, double NewHeight, double OldHeight, Dimension Key)
     { 
       // double NewHeight = ts.getTextHeight();
       double Ratio = OldHeight / NewHeight;
       for(XSLFTextParagraph para : ts.getTextParagraphs() )
        for(XSLFTextRun run : para)
         run.setFontSize(Ratio * run.getFontSize());
       double NewNewHeight = ts.getTextHeight();

       PrintLn(LogFile, "Key (" + Key.width + "," + Key.height + "): old height " + OldHeight + ", new height was " + NewHeight + ", is " + NewNewHeight);
     }

    private static void Resize2(XSLFTextShape ts, Dimension Key)
     {
       PrintLn(LogFile," Trying to resize ( " + Key.width + "," + Key.height + ")...");

       try { 
             Rectangle2D anchor = ts.getAnchor();
             double OldWidth  = anchor.getWidth();
             double OldHeight = anchor.getHeight();
             PrintLn(LogFile,"   old (" + OldWidth + "," + OldHeight + ")");

             anchor = ts.resizeToFitText();
             double NewWidth  = anchor.getWidth();
             double NewHeight = anchor.getHeight();
             PrintLn(LogFile,"   new (" + NewWidth + "," + NewHeight + ")");

             double WRatio = OldWidth  / NewWidth;
             double HRatio = OldHeight / NewHeight;
             double Ratio = (WRatio < HRatio ? WRatio : HRatio);
             int nRun=0;
             for(XSLFTextParagraph para : ts.getTextParagraphs() )
              for(XSLFTextRun run : para)
               { double OldFS = run.getFontSize();
                 double NewFS = Ratio * OldFS;
                 run.setFontSize(NewFS);
                 double NewFS2 = run.getFontSize();
                 nRun++;
                 PrintLn(LogFile,"     run " + nRun + " old, wnew, inew FS = " + OldFS + "," + NewFS + "," + NewFS2);
               };
             anchor = ts.resizeToFitText();
             NewWidth  = anchor.getWidth();
             NewHeight = anchor.getHeight();
             PrintLn(LogFile,"    new new (" + NewWidth + "," + NewHeight + ")");
           } 
       catch(Exception e)
           { 
             PrintLn(LogFile,"** couldn't resize");
           }
     }

    public static Color getColor(PaintStyle ps) 
     {
       if (!(ps instanceof SolidPaint)) return null;
       return DrawPaint.applyColorTransform(((SolidPaint)ps).getSolidColor());
     }

    private static String GetTextRunFormat(XSLFTextRun run)
     { 
       if (!WriteFormats) return new String("");

       String Format = new String();

       PaintStyle.SolidPaint pssp = (PaintStyle.SolidPaint) run.getFontColor();
       Color color = DrawPaint.applyColorTransform(pssp.getSolidColor());
       String ColorString = color.getRed() + "_" + color.getGreen() + "_" + color.getBlue();
       Format = Format + " color " + ColorString;

       Format = Format + " font " + run.getFontFamily().replaceAll(" ","_");
       Format = Format + " size " + run.getFontSize();

       XSLFHyperlink Link = run.getHyperlink();
       if (Link!=null)
        Format = Format + " URL " + Link.getAddress();

       if (run.isBold()) 
        Format=Format + " bold ";
       if (run.isItalic()) 
        Format=Format + " italic ";

       return Format;
     }

    private static void SetTextRunFormat(XSLFTextRun run, String Format)
     { String Tokens[] = Format.split(" ");
       for(int nt=0; nt<Tokens.length; nt++)
        { 
          if ( Tokens[nt].equalsIgnoreCase("color") )
           { nt++;
             //run.setFontColor(DrawPaint.createSolidPaint(Color.getColor(Tokens[nt])));
             String cc[] = Tokens[nt].split("_");
             run.setFontColor(new Color(Integer.parseInt(cc[0]), Integer.parseInt(cc[1]), Integer.parseInt(cc[2])));
           }
          else if ( Tokens[nt].equalsIgnoreCase("font") )
           { nt++;
             run.setFontFamily(Tokens[nt].replaceAll("_"," "));
           }
          else if ( Tokens[nt].equalsIgnoreCase("size") )
           { nt++;
             run.setFontSize(Double.parseDouble(Tokens[nt]));
           }
          else if ( Tokens[nt].equalsIgnoreCase("URL") )
           { nt++;
             String URL=Tokens[nt];
             if (URL.equalsIgnoreCase("none"))
              { //CTTextCharacterProperties p=run.getRPr(false);
                //if (p!=null)
                // p.unsetHlinkClick();
                //XSLFHyperlink Link=run.getHyperlink();
                //if (Link!=null)
                // Link.setAddress(null);
              }
             else
              { XSLFHyperlink Link=run.createHyperlink();
                Link.setAddress(URL);
              };
           }
          else if ( Tokens[nt].equalsIgnoreCase("URL") )
           { nt++;
             run.setFontFamily(Tokens[nt]);
           }
          else if ( Tokens[nt].equalsIgnoreCase("bold") )
           run.setBold(true);
          else if ( Tokens[nt].equalsIgnoreCase("italic") )
           run.setItalic(true);
        };
     }

    private static void SetTextShapeFormat(XSLFTextShape ts, String Format)
     { for(XSLFTextParagraph para : ts.getTextParagraphs() )
        for(XSLFTextRun run : para)
         SetTextRunFormat(run, Format);
     }

    private static String GetTextShapeFormat(XSLFTextShape ts)
     { String Format=null;
       for(XSLFTextParagraph para : ts.getTextParagraphs() )
        for(XSLFTextRun run : para)
         { if (Format==null) 
            Format=GetTextRunFormat(run);
           else if ( !Format.equalsIgnoreCase(GetTextRunFormat(run)))
            return null;
         };
       return Format;
     }

    private static boolean HasHyperlink(XSLFTextShape ts)
     { 
       for(XSLFTextParagraph para : ts.getTextParagraphs() )
        for(XSLFTextRun run : para)
         if (run.getHyperlink() != null)
          return true;

       return false;
     }

    private static boolean IsMulticolored(XSLFTextShape ts)
     { 
       Color FirstColor=null;
       for(XSLFTextParagraph para : ts.getTextParagraphs() )
        for(XSLFTextRun run : para)
          { Color ThisColor = getColor(run.getFontColor());
            if (FirstColor==null)
             FirstColor = ThisColor;
            else if ( !(FirstColor.equals(ThisColor) ) )
             return true;
          }
       return false;
     }

    private static void ProcessTextShape(XSLFTextShape ts)
     {
       String OldText = ts.getText().toString();
       if ( WideOnly && !HasWideCharacters(OldText) )
        return;

       nText++;
       TotalTextShapes++;
       Dimension Key = new Dimension();
       Key.setSize(nText, 0);
       boolean Changed  = false;

       if ( Mode==ModeValue.EXTRACT )
        {
          Print(TextFile,"SHAPE " + TotalTextShapes);
          if (HasHyperlink(ts))
           Print(TextFile," HYPERLINK ");
          if (IsMulticolored(ts))
           Print(TextFile," MULTICOLORED ");
          Print(TextFile," \n");

          String Format=GetTextShapeFormat(ts);
          if (Format==null) Format="";

          PrintLn(TextFile,TEXT_KEYWORD + " " + nText + " 0" + Format);
          PrintLn(TextFile,TEXT_SEPARATOR);
          PrintLn(TextFile,OldText);
          PrintLn(TextFile,TEXT_SEPARATOR + "\n");
        }
       else if (Translations.containsKey(Key))
        { if (LogFile!=null)
           PrintLn(LogFile," ** Replacing (" + Key.width + "," + Key.height + ")");
          TextShapesReplaced++; 
          ts.setText(Translations.get(Key) );
          if (Formats.containsKey(Key))
           SetTextShapeFormat(ts, Formats.get(Key));
          Translations.remove(Key);  
          Changed = true;
          if (Mode==ModeValue.TRANSLATE && Changed && Autosize)
           Resize2(ts, Key);
          return;
        };

       if (OmitRuns) return;

       int nRun=0;
       for(XSLFTextParagraph para : ts.getTextParagraphs() )
        { List<XSLFTextRun> EmptyTextRuns = new ArrayList<XSLFTextRun>();
          for(XSLFTextRun run : para)
           { 
             nRun++;
             TotalTextRuns++;
             Key.setSize(nText, nRun);

             if ( Mode==ModeValue.EXTRACT )
              { 
                String Format = GetTextRunFormat(run);
                PrintLn(TextFile,TEXT_KEYWORD + " " + Key.width + " " + Key.height + Format);
                PrintLn(TextFile,TEXT_SEPARATOR);
                PrintLn(TextFile,run.getRawText().toString());
                PrintLn(TextFile,TEXT_SEPARATOR + "\n");
              }
             else if (Translations.containsKey(Key))
              { String NewText = Translations.get(Key);
                if ( NewText==null || NewText.trim().length()==0 )
                 { 
                   PrintLn(LogFile," ** Removing (" + Key.width + "," + Key.height + ")");
                   run.setText("");
                   EmptyTextRuns.add(run);
                 }
                else
                 { 
                   PrintLn(LogFile," ** Replacing (" + Key.width + "," + Key.height + ")");
                   TextRunsReplaced++; 
                   // remove trailing CR from text runs
                   run.setText(NewText.replace("\n"," ").replace("\r"," "));
                   if (Formats.containsKey(Key))
                    SetTextRunFormat(run, Formats.get(Key));
                 }
                Translations.remove(Key);
                Changed = true;
              }
           }
          para.getTextRuns().removeAll(EmptyTextRuns);
        };

       if (Mode==ModeValue.TRANSLATE && Changed && Autosize)
        Resize2(ts, Key);

     }
   
    private static void extractText(XSLFShapeContainer data, boolean skipPlaceholders, StringBuilder text) 
    {
     for (XSLFShape s : data) 
      {
         if (s instanceof XSLFShapeContainer) 
          {
             extractText((XSLFShapeContainer)s, skipPlaceholders, text);
          } 
         else if (s instanceof XSLFTextShape) 
          {
             XSLFTextShape ts = (XSLFTextShape)s;

             // Skip non-customised placeholder text
             if (skipPlaceholders && ts.isPlaceholder())
              continue;

            ProcessTextShape(ts);

          } 
         else if (s instanceof XSLFTable) 
          {
             XSLFTable ts = (XSLFTable)s;

             // Skip non-customised placeholder text
             for (XSLFTableRow r : ts) 
              for (XSLFTableCell c : r) 
               { 
                 TotalTableEntries++;
                 ProcessTextShape(c);
               }
         } // if (s instance of ...)

       } // for (XSLFShape s : data) 

    } // private static void extractText(XSLFShapeContainer data, boolean skipPlaceholders, StringBuilder text) 

} // public class TranslatePPTX extends POIXMLTextExtractor {
