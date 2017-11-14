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
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFCommentAuthors;
import org.apache.poi.xslf.usermodel.XSLFComments;
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
        public static boolean Verbose=false;
        public static boolean WideOnly=false;
        public static FileWriter LogFile=null;
        public static FileWriter TextFile=null;
        public static int TotalTextShapes=0;
        public static int TotalTextRuns=0;
        public static int TotalTableEntries=0;

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
	   System.err.println("  --WriteLog");
	   System.err.println("  --WideOnly");
	   System.err.println("  --Verbose");
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
                 while ( (Line = br.readLine()) != null)
                  { 
                    LineNum++;
                    
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

                    // sanity checks:
                    //  (a) key (NT, NR) does not already exist in table
                    //  (b) for a given NT, *either* (NT,0) *or* (NT,i>0) may exist,
                    //      but not both (each TextShape is either translated all-at-once
                    //      *or* by individual TextRuns, but not both
                    if (Translations.containsKey(Key))
                     ErrExit(FileName + ":" + LineNum + ":" + "Key (" + Key.width + "," + Key.height + ") exists twice");
                    if (Key.height>0)
                     { Dimension KeyPrime=new Dimension(Key);
                       KeyPrime.height=0;
                       if (Translations.containsKey(KeyPrime))
                        ErrExit(FileName + ":" + LineNum + ":" + "Key(" + Key.width + "," + Key.height + ") conflicts with existing (" + Key.width + ",0)");
                     }

                    // next line must be text separator
                    if ( ! br.readLine().equals(TEXT_SEPARATOR) )
                     ErrExit(FileName + ":" + LineNum + ":" + "expected separator string (=================)");
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

               }
           catch(IOException e)
               { 
                  ErrExit("could not open file " + FileName);
               }
         };

	public static void PrintLn(FileWriter f, String s)
	 {  
           if (f==null) return;
           try 
            { 
               f.write(s + "\n");
               f.flush();
            }
           catch(Exception e) 
            {
            }
         }

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
           for(int narg=1; narg<args.length; narg++)
            { 
              if ( args[narg].equalsIgnoreCase("--Verbose") )
               Verbose=true;
              else if ( args[narg].equalsIgnoreCase("--WriteLog") )
               WriteLog=true;
              else if ( args[narg].equalsIgnoreCase("--WideOnly") )
               WideOnly=true;
              else if ( args[narg].equalsIgnoreCase("--Translations") )
               { Mode=ModeValue.TRANSLATE;
                 ReadTranslations(args[narg+1]);
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
            TextFile = new FileWriter(FileBase + ".text");

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
              String OutFileName = new String(FileBase + "_Translated.pptx");
              try (FileOutputStream OutFile = new FileOutputStream(OutFileName) )
               {
                 ppt.write(OutFile);
                 System.out.println("Wrote translated document to " + OutFileName + ".");
               }
              catch(Exception e)
               {
                 ErrExit("could not write output file " + OutFileName);
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
           System.out.println(" Handling slide " + nSlide);

          StringBuilder SlideText = new StringBuilder();
          SlideText.append(getText(slide, slideText, notesText, masterText));
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

             // Skip strings without Japanese text
             String OldText = ts.getText().toString();
             if ( WideOnly && !HasWideCharacters(OldText) )
              continue;

             nText++;
             TotalTextShapes++;

             double OldHeight = ts.getTextHeight();
             boolean Changed  = false;
             Dimension Key = new Dimension();
             Key.setSize(nText, 0);

             if ( Mode==ModeValue.EXTRACT )
              {
                PrintLn(TextFile,TEXT_KEYWORD + " " + nText + " 0");
                PrintLn(TextFile,TEXT_SEPARATOR);
                PrintLn(TextFile,OldText);
                PrintLn(TextFile,TEXT_SEPARATOR + "\n");
              }
             else if (Translations.containsKey(Key))
              { if (LogFile!=null)
                 PrintLn(LogFile," ** Replacing (" + Key.width + "," + Key.height + ")");
                ts.setText(Translations.get(Key) );
                Translations.remove(Key);
                Changed = true;
              };

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
                      PrintLn(TextFile,TEXT_KEYWORD + " " + Key.width + " " + Key.height);
                      PrintLn(TextFile,TEXT_SEPARATOR);
                      PrintLn(TextFile,run.getRawText().toString());
                      PrintLn(TextFile,TEXT_SEPARATOR + "\n");
                    }
                   else if (Translations.containsKey(Key))
                    { String NewText = Translations.get(Key);
                      if ( NewText.trim().length()==0 )
                       { 
                         PrintLn(LogFile," ** Removing (" + Key.width + "," + Key.height + ")");
                         EmptyTextRuns.add(run);
                       }
                      else
                       { 
                         PrintLn(LogFile," ** Replacing (" + Key.width + "," + Key.height + ")");
                         run.setText(NewText);
                       }
                      Translations.remove(Key);
                      Changed = true;
                    }
                 }
                para.getTextRuns().removeAll(EmptyTextRuns);
             }

             if (Mode==ModeValue.TRANSLATE && Changed)
              { double NewHeight = ts.getTextHeight();
                Resize(ts, NewHeight, OldHeight, Key);
              };
          } 
         else if (s instanceof XSLFTable) 
          {
             XSLFTable ts = (XSLFTable)s;

             // Skip non-customised placeholder text
             for (XSLFTableRow r : ts) 
              for (XSLFTableCell c : r) 
               { 
                 String OldText = c.getText().toString();
                 if ( WideOnly && !HasWideCharacters(OldText) )
                  continue;

                 nText++;
                 TotalTableEntries++;
                 Dimension Key = new Dimension();
                 Key.setSize(nText, 0);
                 if (Mode==ModeValue.EXTRACT)
                  { PrintLn(TextFile,"\n" + TEXT_KEYWORD + " " + nText + " " + "0 (table)");
                    PrintLn(TextFile,TEXT_SEPARATOR);
                    PrintLn(TextFile,c.getText().toString());
                    PrintLn(TextFile,TEXT_SEPARATOR + "\n");
                  }
                 else if ( Translations.containsKey(Key) )
                  {
                    XSLFTextShape cts = (XSLFTextShape)c;
                    PrintLn(LogFile," ** Replacing (" + Key.width + "," + Key.height + ")");

                    double OldWidth=0.0;
                    double OldHeight=0.0;
                    try {

                           Rectangle2D anchor = cts.getAnchor();
                           OldWidth  = anchor.getWidth();
                           OldHeight = anchor.getHeight();
                           PrintLn(LogFile,"    old size (" + OldWidth + "," + OldHeight + ")");
                        }
                    catch(Exception e)
                        {
                           PrintLn(LogFile,"    failed to get old anchor");
                        }

                    try {
                          cts.setText(Translations.get(Key));
                        }
                    catch(Exception e) 
                        { 
                          System.err.println("** FAILED to replace (" + Key.width + "," + Key.height + ")");
                          PrintLn(LogFile,"** FAILED to replace (" + Key.width + "," + Key.height + ")");
                        }

                    Translations.remove(Key);

                    if (OldHeight!=0.0)
                     { try { 
                             Rectangle2D anchor = cts.resizeToFitText();
                             double NewWidth  = anchor.getWidth();
                             double NewHeight = anchor.getHeight();
                             PrintLn(LogFile,"    new size (" + NewWidth + "," + NewHeight + ")");
                             double WRatio = OldWidth  / NewWidth;
                             double HRatio = OldHeight / NewHeight;
                             double Ratio = (WRatio < HRatio ? WRatio : HRatio);
/*
                             for(XSLFTextParagraph para : cts.getTextParagraphs() )
                              for(XSLFTextRun run : para)
                               run.setFontSize(Ratio * run.getFontSize());
*/
                             int nRun=0;
                             for(XSLFTextParagraph para : cts.getTextParagraphs() )
                              for(XSLFTextRun run : para)
                               { double OldFS = run.getFontSize();
                                 double NewFS = Ratio * OldFS;
                                 run.setFontSize(NewFS);
                                 double NewFS2 = run.getFontSize();
                                 nRun++;
                                 PrintLn(LogFile,"   run " + nRun + " old, wnew, inew FS = " + OldFS + "," + NewFS + "," + NewFS2);
                               };
                             anchor = cts.resizeToFitText();
                             NewWidth  = anchor.getWidth();
                             NewHeight = anchor.getHeight();
                             PrintLn(LogFile,"    new new size (" + NewWidth + "," + NewHeight + ")");
                           }
                       catch(Exception e)
                           { 
                             PrintLn(LogFile,"** couldn't resize");
                           }
                     }
                  }
               }
         } // if (s instance of ...)

       } // for (XSLFShape s : data) 

    } // private static void extractText(XSLFShapeContainer data, boolean skipPlaceholders, StringBuilder text) 

} // public class TranslatePPTX extends POIXMLTextExtractor {
