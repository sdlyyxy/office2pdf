using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Windows.Forms;
using System.Web;

namespace HelloWorld{
	class Program{
		static void word2pdf(string sourcePath,string targetPath){
			// Console.WriteLine("hello..");
			Word.Application myWordApp;
			Word.Document myWordDoc;
			myWordApp = new Word.ApplicationClass();
			// object filepath =  fileString;
			object filepath =  sourcePath;
			object oMissing = System.Reflection.Missing.Value;
			myWordDoc = myWordApp.Documents.Open(ref filepath, ref oMissing, ref oMissing, ref oMissing, ref oMissing,ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            Word.WdExportFormat paramExportFormat = Word.WdExportFormat.wdExportFormatPDF;
            bool paramOpenAfterExport = false;
            Word.WdExportOptimizeFor paramExportOptimizeFor =
                    // Word.WdExportOptimizeFor.wdExportOptimizeForPrint;
                    Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen;
            Word.WdExportRange paramExportRange = Word.WdExportRange.wdExportAllDocument;
            int paramStartPage = 0;
            int paramEndPage = 0;
            Word.WdExportItem paramExportItem = Word.WdExportItem.wdExportDocumentContent;
            bool paramIncludeDocProps = true;
            bool paramKeepIRM = true;
            Word.WdExportCreateBookmarks paramCreateBookmarks =
                    Word.WdExportCreateBookmarks.wdExportCreateWordBookmarks;
            bool paramDocStructureTags = true;
            bool paramBitmapMissingFonts = true;
            bool paramUseISO19005_1 = false;
			string paramExportFilePath=targetPath;
			myWordDoc.ExportAsFixedFormat(paramExportFilePath,
					paramExportFormat, paramOpenAfterExport,
					paramExportOptimizeFor, paramExportRange, paramStartPage,
					paramEndPage, paramExportItem, paramIncludeDocProps,
					paramKeepIRM, paramCreateBookmarks, paramDocStructureTags,
					paramBitmapMissingFonts, paramUseISO19005_1,
					ref oMissing);
            myWordDoc.Close(ref oMissing, ref oMissing, ref oMissing);
            myWordDoc = null;
            myWordApp.Quit(ref oMissing, ref oMissing, ref oMissing);
            myWordDoc = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
		}
		static void ppt2pdf(string sourcePath,string targetPath){
			bool result=false;
			object missing = Type.Missing;
			PowerPoint.ApplicationClass application = null;
			PowerPoint.Presentation persentation = null;
			try
			{
				application = new PowerPoint.ApplicationClass();
				persentation = application.Presentations.Open(sourcePath, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
				persentation.SaveAs(targetPath,PowerPoint.PpSaveAsFileType.ppSaveAsPDF, Microsoft.Office.Core.MsoTriState.msoTrue);
				result = true;
			}
			catch
			{
				result = false;
			}
			finally
			{
				if (persentation != null)
				{
					persentation.Close();
					persentation = null;
				}
				if (application != null)
				{
					application.Quit();
					application = null;
				}
				GC.Collect();
				GC.WaitForPendingFinalizers();
				GC.Collect();
				GC.WaitForPendingFinalizers();
			}
			// return result;
		}
		static void Main(string[] args){
			string _directory=Directory.GetCurrentDirectory();
			Console.WriteLine(_directory);
			string[] files=Directory.GetFiles(@".","*.doc*");
			foreach(string file in files){
				// int pos=file.LastIndexOf('.');
				string sourcePath=file;
				if(sourcePath[sourcePath.Length-1]=='f')continue;
				string _file=_directory+@"\"+file;
				// _file.Remove(pos);
				_file+=@".pdf";
				Console.WriteLine(Directory.Exists(_file));
				Console.WriteLine(_file);
				if(Directory.Exists(_file))continue;
				// Console.Write("Transform ");
				Console.WriteLine(sourcePath);
				// Console.WriteLine(" ?");
				// char c=(char)Console.Read();
				// if(c!='y')continue;
				word2pdf(_directory+@"\"+sourcePath,_file);
			}
			files=Directory.GetFiles(@".","*.ppt*");
			foreach(string file in files){
				// int pos=file.LastIndexOf('.');
				string sourcePath=file;
				string _file=_directory+@"\"+file;
				if(sourcePath[sourcePath.Length-1]=='f')continue;				
				// _file.Remove(pos);
				_file+=@".pdf";
				// Console.WriteLine(Directory.Exists(_file));
				Console.WriteLine(_file);
				if(Directory.Exists(_file))continue;
				// Console.Write("Transform ");
				Console.WriteLine(sourcePath);
				// Console.WriteLine(" ?");
				// char c=(char)Console.Read();
				// if(c!='y')continue;
				ppt2pdf(_directory+@"\"+sourcePath,_file);
			}
		}
	}
}
//Microsoft.Office.Interop.PowerPoint._Presentation.ExportAsFixedFormat(string, Microsoft.Office.Interop.PowerPoint.PpFixedFormatType, Microsoft.Office.Interop.PowerPoint.PpFixedFormatIntent, Microsoft.Office.Core.MsoTriState, Microsoft.Office.Interop.PowerPoint.PpPrintHandoutOrder, Microsoft.Office.Interop.PowerPoint.PpPrintOutputType, Microsoft.Office.Core.MsoTriState, Microsoft.Office.Interop.PowerPoint.PrintRange, Microsoft.Office.Interop.PowerPoint.PpPrintRangeType, string, bool, bool, bool, bool, bool, object)”最匹配的重载方法具有一些无效参�?
/*


/* use saveasfixedformat
			// object oMissing = System.Reflection.Missing.Value;
			// PowerPoint.Application myPowerPointApp;
			// PowerPoint.Presentation myPowerPointPresentation;
			// myPowerPointApp = new PowerPoint.ApplicationClass();
			// myPowerPointPresentation = myPowerPointApp.Presentations.Open(sourcePath, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
            // PowerPoint.PpFixedFormatType paramFixedFormatType = PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF;
			// PowerPoint.PpFixedFormatIntent paramFixedFormatIntent=PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentPrint;
            // Microsoft.Office.Core.MsoTriState paramFrameSlides=Microsoft.Office.Core.MsoTriState.msoFalse;
            // PowerPoint.PpPrintHandoutOrder paramHandoutOrder=PowerPoint.PpPrintHandoutOrder.ppPrintHandoutVerticalFirst;
            // PowerPoint.PpPrintOutputType paramOutputType=PowerPoint.PpPrintOutputType.ppPrintOutputSlides;
            // Microsoft.Office.Core.MsoTriState paramPrintHidenSlides=Microsoft.Office.Core.MsoTriState.msoFalse;
            // PowerPoint.PrintRange paramPrintRange=PowerPoint.PrintRange.Start();
			// myPowerPointPresentation.ExportAsFixedFormat(targetPath,paramFixedFormatType, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing,oMissing);
			// myPowerPointPresentation.ExportAsFixedFormat(targetPath,paramFixedFormatType, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing,oMissing);
            // myPowerPointPresentation.Close();
            // myPowerPointPresentation = null;
            // myPowerPointApp.Quit();
			// myPowerPointApp=null;
            // GC.Collect();
            // GC.WaitForPendingFinalizers();
            // GC.Collect();
            // GC.WaitForPendingFinalizers();
*/
/*
//将word文档转换成PDF格式
    private bool Convert(string sourcePath, string targetPath, Word.WdExportFormat exportFormat)
    {
        bool result;
        object paramMissing = Type.Missing;
        Word.ApplicationClass wordApplication = new Word.ApplicationClass();
        Word.Document wordDocument = null;
        try
        {
            object paramSourceDocPath = sourcePath;
            string paramExportFilePath = targetPath;

            Word.WdExportFormat paramExportFormat = exportFormat;
            bool paramOpenAfterExport = false;
            Word.WdExportOptimizeFor paramExportOptimizeFor =
                    Word.WdExportOptimizeFor.wdExportOptimizeForPrint;
            Word.WdExportRange paramExportRange = Word.WdExportRange.wdExportAllDocument;
            int paramStartPage = 0;
            int paramEndPage = 0;
            Word.WdExportItem paramExportItem = Word.WdExportItem.wdExportDocumentContent;
            bool paramIncludeDocProps = true;
            bool paramKeepIRM = true;
            Word.WdExportCreateBookmarks paramCreateBookmarks =
                    Word.WdExportCreateBookmarks.wdExportCreateWordBookmarks;
            bool paramDocStructureTags = true;
            bool paramBitmapMissingFonts = true;
            bool paramUseISO19005_1 = false;

            wordDocument = wordApplication.Documents.Open(
                    ref paramSourceDocPath, ref paramMissing, ref paramMissing,
                    ref paramMissing, ref paramMissing, ref paramMissing,
                    ref paramMissing, ref paramMissing, ref paramMissing,
                    ref paramMissing, ref paramMissing, ref paramMissing,
                    ref paramMissing, ref paramMissing, ref paramMissing,
                    ref paramMissing);

            if (wordDocument != null)
                wordDocument.ExportAsFixedFormat(paramExportFilePath,
                        paramExportFormat, paramOpenAfterExport,
                        paramExportOptimizeFor, paramExportRange, paramStartPage,
                        paramEndPage, paramExportItem, paramIncludeDocProps,
                        paramKeepIRM, paramCreateBookmarks, paramDocStructureTags,
                        paramBitmapMissingFonts, paramUseISO19005_1,
                        ref paramMissing);
            result = true;
        }
        finally
        {
            if (wordDocument != null)
            {
                wordDocument.Close(ref paramMissing, ref paramMissing, ref paramMissing);
                wordDocument = null;
            }
            if (wordApplication != null)
            {
                wordApplication.Quit(ref paramMissing, ref paramMissing, ref paramMissing);
                wordApplication = null;
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        return result;
    }

    //将excel文档转换成PDF格式
    private bool Convert(string sourcePath, string targetPath, XlFixedFormatType targetType)
    {
        bool result;
        object missing = Type.Missing;
        Excel.ApplicationClass application = null;
        Workbook workBook = null;
        try
        {
            application = new Excel.ApplicationClass();
            object target = targetPath;
            object type = targetType;
            workBook = application.Workbooks.Open(sourcePath, missing, missing, missing, missing, missing,
                    missing, missing, missing, missing, missing, missing, missing, missing, missing);

            workBook.ExportAsFixedFormat(targetType, target, XlFixedFormatQuality.xlQualityStandard, true, false, missing, missing, missing, missing);
            result = true;
        }
        catch
        {
            result = false;
        }
        finally
        {
            if (workBook != null)
            {
                workBook.Close(true, missing, missing);
                workBook = null;
            }
            if (application != null)
            {
                application.Quit();
                application = null;
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        return result;
    }

    //将ppt文档转换成PDF格式
    private bool Convert(string sourcePath, string targetPath, PpSaveAsFileType targetFileType)
    {
        bool result;
        object missing = Type.Missing;
        PowerPoint.ApplicationClass application = null;
        Presentation persentation = null;
        try
        {
            application = new PowerPoint.ApplicationClass();
            persentation = application.Presentations.Open(sourcePath, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
            persentation.SaveAs(targetPath, targetFileType, Microsoft.Office.Core.MsoTriState.msoTrue);

            result = true;
        }
        catch
        {
            result = false;
        }
        finally
        {
            if (persentation != null)
            {
                persentation.Close();
                persentation = null;
            }
            if (application != null)
            {
                application.Quit();
                application = null;
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        return result;
    }
	
	*/