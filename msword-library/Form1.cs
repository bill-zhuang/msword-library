using System;
using System.Windows.Forms;
using System.IO;

using MSWord = Microsoft.Office.Interop.Word;
using System.Collections.Generic;

using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace wordLibrary
{
    public partial class formWordLibrary : Form
    {
        public string docFolderPath;
        public string saveFolderPath;
        public string docProcessPath;

        public formWordLibrary()
        {
            docFolderPath = "";
            saveFolderPath = "";
            docProcessPath = "";

            InitializeComponent();
        }

        private void btnOpenFolder_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    docFolderPath = folderDialog.SelectedPath;
                    //create save folder
                    saveFolderPath = docFolderPath + "/save/";
                    createDirectory(saveFolderPath);
                    //
                    textFolderPath.Text = docFolderPath;
                }
            }
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "doc/docx files(*.doc;*.docx)|*.doc;*.docx|doc files(*.doc)|*.doc|docx files(*.docx)|*.docx";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                docProcessPath = fileDialog.FileName.ToString();
                //create save folder
                saveFolderPath = Path.GetDirectoryName(docProcessPath) + "/save/";
                createDirectory(saveFolderPath);
                //
                textFolderPath.Text = docProcessPath;
            }
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //disable button
            btnOpenFolder.Enabled = false;
            btnSelectFile.Enabled = false;
            btnRun.Enabled = false;
            //clear text box
            textBoxProcess.Text = string.Empty;
            if (docFolderPath != "")
            {
                //
            }
            else if (docProcessPath != "")
            {
                if (File.Exists(docProcessPath))
                {
                    //int pageNum = getDocPageNum(docProcessPath);
                    //textBoxProcess.AppendText(docProcessPath + "页数：" + pageNum + "\r\n");
                    //List<string> docContent = getDocText(docProcessPath);
                    //textBoxProcess.AppendText(docProcessPath + "内容：" + string.Join("\r\n", docContent.ToArray()) + "\r\n");
                    insertTextWaterMark(docProcessPath, "test", "宋体", 15.0f);
                    MessageBox.Show("处理结束。");
                }
                else
                {
                    MessageBox.Show(docProcessPath + "文件不存在");
                }
            }
            else
            {
                MessageBox.Show("请选择要处理的文件夹或者文件！");
            }

            //enable button
            btnOpenFolder.Enabled = true;
            btnSelectFile.Enabled = true;
            btnRun.Enabled = true;
        }

        //test pass
        public int getDocPageNum(string docPath)
        {
            int pageNum = 0;

            if (File.Exists(docPath))
            {
                try
                {
                    MSWord._Application wordApp = new MSWord.Application() { Visible = false, DisplayAlerts = MSWord.WdAlertLevel.wdAlertsNone };
                    MSWord._Document doc = wordApp.Documents.OpenNoRepairDialog(docPath);
                    if (doc != null)
                    {
                        object missing = System.Reflection.Missing.Value;
                        pageNum = doc.ComputeStatistics(MSWord.WdStatistic.wdStatisticPages, ref missing);

                        doc.Close(MSWord.WdSaveOptions.wdDoNotSaveChanges, MSWord.WdOriginalFormat.wdOriginalDocumentFormat, false);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);

                        wordApp.Quit(MSWord.WdSaveOptions.wdDoNotSaveChanges, MSWord.WdOriginalFormat.wdOriginalDocumentFormat, false);
                    }
                }
                catch (Exception e)
                {
                    textBoxProcess.AppendText(docProcessPath + "出错了：" + e.Message + "\r\n");
                    killMSWordProcess();
                }
            }

            return pageNum;
        }

        //under test
        public void getDocLastPage(string docPath)
        {
            try
            {
                MSWord._Application wordApp = new MSWord.Application() { Visible = false, DisplayAlerts = MSWord.WdAlertLevel.wdAlertsNone };
                MSWord._Document doc = wordApp.Documents.OpenNoRepairDialog(docPath);
                if (doc != null)
                {
                    object missing = System.Reflection.Missing.Value;
                    object oToName = "1";
                    object oToWhich = MSWord.WdGoToDirection.wdGoToNext;
                    object oToWhat = MSWord.WdGoToItem.wdGoToPage;
                    MSWord.Document newDoc = wordApp.Documents.Add(ref missing, ref missing, ref missing, ref missing);
                    MSWord.Range rngPageStart = null;
                    MSWord.Range rngPagePrevious = doc.Content;
                    doc.Activate();
                    //Start with first page
                    rngPageStart = wordApp.Selection.GoTo(ref oToWhat, ref missing, ref missing, ref oToName);
                    while (!rngPagePrevious.InRange(rngPageStart))
                    {
                        //CopyPage(wdApp.Selection, newDoc);
                        rngPagePrevious = rngPageStart;
                        rngPageStart = wordApp.Selection.GoTo(ref oToWhat, ref oToWhich, ref missing, ref missing);
                    }

                    doc.Close(MSWord.WdSaveOptions.wdDoNotSaveChanges, MSWord.WdOriginalFormat.wdOriginalDocumentFormat, false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);

                    wordApp.Quit(MSWord.WdSaveOptions.wdDoNotSaveChanges, MSWord.WdOriginalFormat.wdOriginalDocumentFormat, false);
                }
            }
            catch (Exception e)
            {
                textBoxProcess.AppendText(docProcessPath + "出错了：" + e.Message + "\r\n");
                killMSWordProcess();
            }
        }

        //test pass
        public List<string> getDocText(string docPath)
        {
            List<string> docContent = new List<string>();
            if (File.Exists(docPath))
            {
                try
                {
                    MSWord._Application wordApp = new MSWord.Application() { Visible = false, DisplayAlerts = MSWord.WdAlertLevel.wdAlertsNone };
                    MSWord._Document doc = wordApp.Documents.OpenNoRepairDialog(docPath);
                    if (doc != null)
                    {
                        //String read = string.Empty;
                        for (int i = 0; i < doc.Paragraphs.Count; i++)
                        {
                            string lineContent = doc.Paragraphs[i + 1].Range.Text.Trim();
                            if (lineContent != string.Empty)
                            {
                                docContent.Add(lineContent);
                            }
                        }

                        doc.Close(MSWord.WdSaveOptions.wdDoNotSaveChanges, MSWord.WdOriginalFormat.wdOriginalDocumentFormat, false);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);

                        wordApp.Quit(MSWord.WdSaveOptions.wdDoNotSaveChanges, MSWord.WdOriginalFormat.wdOriginalDocumentFormat, false);
                    }
                }
                catch (Exception e)
                {
                    textBoxProcess.AppendText(docProcessPath + "出错了：" + e.Message + "\r\n");
                    killMSWordProcess();
                }
            }

            return docContent;
        }

        public void removeDocHeadeAndFooter(string docPath)
        {
        }

        //test pass
        //How to: Remove the headers and footers from a word processing document (Open XML SDK)
        //https://msdn.microsoft.com/EN-US/library/office/hh181053.aspx
        public void removeDocxHeaderAndFooter(string docxPath)
        {
            // Given a document name, remove all of the headers and footers
            // from the document.
            using (WordprocessingDocument docx = WordprocessingDocument.Open(docxPath, true))
            {
                // Get a reference to the main document part.
                var docxPart = docx.MainDocumentPart;

                //if (docPart.ImageParts.Count() > 0)
                //{
                //    docPart.DeleteParts(docPart.ImageParts);
                //}

                // Count the header and footer parts and continue if there are any.
                if (docxPart.HeaderParts.Count() > 0 || docxPart.FooterParts.Count() > 0)
                {
                    // Remove the header and footer parts.
                    docxPart.DeleteParts(docxPart.HeaderParts);
                    docxPart.DeleteParts(docxPart.FooterParts);

                    // Get a reference to the root element of the main document part.
                    Document document = docxPart.Document;

                    // Remove all references to the headers and footers.

                    // First, create a list of all descendants of type
                    // HeaderReference. Then, navigate the list and call
                    // Remove on each item to delete the reference.
                    var headers = document.Descendants<HeaderReference>().ToList();
                    foreach (var header in headers)
                    {
                        header.Remove();
                    }

                    // First, create a list of all descendants of type
                    // FooterReference. Then, navigate the list and call
                    // Remove on each item to delete the reference.
                    var footers = document.Descendants<FooterReference>().ToList();
                    foreach (var footer in footers)
                    {
                        footer.Remove();
                    }

                    // Save the changes.
                    document.Save();
                }
            }
        }

        //under test
        public void insertPicture2Header(string docPath, string logoPath)
        {
            object missing = System.Reflection.Missing.Value;
            object oFalse = false;
            object oTrue = true;
            MSWord._Application wordApp = new MSWord.Application() { Visible = false, DisplayAlerts = MSWord.WdAlertLevel.wdAlertsNone };
            MSWord._Document doc = wordApp.Documents.OpenNoRepairDialog(docPath);
            if (doc != null)
            {
                //EMBEDDING LOGOS IN THE DOCUMENT
                //SETTING FOCUES ON THE PAGE HEADER TO EMBED THE WATERMARK
                wordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekCurrentPageHeader;

                //THE LOGO IS ASSIGNED TO A SHAPE OBJECT SO THAT WE CAN USE ALL THE
                //SHAPE FORMATTING OPTIONS PRESENT FOR THE SHAPE OBJECT
                MSWord.Shape logoCustom = null;

                //THE PATH OF THE LOGO FILE TO BE EMBEDDED IN THE HEADER
                logoCustom = wordApp.Selection.HeaderFooter.Shapes.AddPicture(logoPath,
                    ref oFalse, ref oTrue, ref missing, ref missing, ref missing, ref missing, ref missing);

                logoCustom.Select(ref missing);
                logoCustom.Name = "CustomLogo";
                logoCustom.Left = (float)MSWord.WdShapePosition.wdShapeLeft;

                //SETTING FOCUES BACK TO DOCUMENT
                wordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekMainDocument;

                doc.Close(MSWord.WdSaveOptions.wdDoNotSaveChanges, MSWord.WdOriginalFormat.wdOriginalDocumentFormat, false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);

                wordApp.Quit(MSWord.WdSaveOptions.wdDoNotSaveChanges, MSWord.WdOriginalFormat.wdOriginalDocumentFormat, false);
            }
        }

        //under test
        //http://www.c-sharpcorner.com/UploadFile/amrish_deep/WordAutomation05102007223934PM/WordAutomation.aspx
        public void insertTextWaterMark(string docPath, string textWaterMark, string fontName, float fontSize, float rotate = 0.0f, float left = 0.0f, float top = 0.0f)
        {
            object missing = System.Reflection.Missing.Value;
            MSWord._Application wordApp = new MSWord.Application() { Visible = false, DisplayAlerts = MSWord.WdAlertLevel.wdAlertsNone };
            MSWord._Document doc = wordApp.Documents.OpenNoRepairDialog(docPath);
            if (doc != null)
            {
                //THE LOGO IS ASSIGNED TO A SHAPE OBJECT SO THAT WE CAN USE ALL THE
                //SHAPE FORMATTING OPTIONS PRESENT FOR THE SHAPE OBJECT
                MSWord.Shape textWatermark = null;

                //INCLUDING THE TEXT WATER MARK TO THE DOCUMENT
                textWatermark = wordApp.Selection.HeaderFooter.Shapes.AddTextEffect(
                    Microsoft.Office.Core.MsoPresetTextEffect.msoTextEffect1,
                    textWaterMark, fontName, fontSize,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    left, top, ref missing);

                textWatermark.Select(ref missing);
                textWatermark.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                textWatermark.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                textWatermark.Fill.Solid();
                textWatermark.Fill.ForeColor.RGB = (Int32)MSWord.WdColor.wdColorGray30;
                textWatermark.Rotation = rotate;
                textWatermark.RelativeHorizontalPosition = MSWord.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin;
                textWatermark.RelativeVerticalPosition = MSWord.WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin;
                textWatermark.Left = (float)MSWord.WdShapePosition.wdShapeCenter;
                textWatermark.Top = (float)MSWord.WdShapePosition.wdShapeCenter;
                textWatermark.Height = wordApp.InchesToPoints(2.4f);
                textWatermark.Width = wordApp.InchesToPoints(6f);

                //SETTING FOCUES BACK TO DOCUMENT
                wordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekMainDocument;

                doc.SaveAs2(saveFolderPath + Path.GetFileName(docPath), MSWord.WdSaveFormat.wdFormatDocument);
                doc.Close(MSWord.WdSaveOptions.wdDoNotSaveChanges, MSWord.WdOriginalFormat.wdOriginalDocumentFormat, false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);

                wordApp.Quit(MSWord.WdSaveOptions.wdDoNotSaveChanges, MSWord.WdOriginalFormat.wdOriginalDocumentFormat, false);
            }
        }

        //under test
        public void insertFooterText(string docPath, string textWaterMark)
        {
            object missing = System.Reflection.Missing.Value;
            MSWord._Application wordApp = new MSWord.Application() { Visible = false, DisplayAlerts = MSWord.WdAlertLevel.wdAlertsNone };
            MSWord._Document doc = wordApp.Documents.OpenNoRepairDialog(docPath);
            if (doc != null)
            {
                //SETTING THE FOCUES ON THE PAGE FOOTER
                wordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekCurrentPageFooter;

                //ENTERING A PARAGRAPH BREAK "ENTER"
                wordApp.Selection.TypeParagraph();

                //INSERTING THE PAGE NUMBERS CENTRALLY ALIGNED IN THE PAGE FOOTER
                wordApp.Selection.Paragraphs.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                wordApp.ActiveWindow.Selection.Font.Name = "Arial";
                wordApp.ActiveWindow.Selection.Font.Size = 8;

                //INSERTING WATERMARK TEXT
                wordApp.ActiveWindow.Selection.TypeText(textWaterMark);

                //wordApp.ActiveWindow.Selection.TypeText("Page ");
                ////get current page
                //Object CurrentPage = MSWord.WdFieldType.wdFieldPage;
                //wordApp.ActiveWindow.Selection.Fields.Add(wordApp.Selection.Range, ref CurrentPage, ref missing, ref missing);
                //wordApp.ActiveWindow.Selection.TypeText(" of ");
                ////get total page
                //Object TotalPages = MSWord.WdFieldType.wdFieldNumPages;
                //wordApp.ActiveWindow.Selection.Fields.Add(wordApp.Selection.Range, ref TotalPages, ref missing, ref missing);

                //SETTING FOCUES BACK TO DOCUMENT
                wordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekMainDocument;

                doc.Close(MSWord.WdSaveOptions.wdDoNotSaveChanges, MSWord.WdOriginalFormat.wdOriginalDocumentFormat, false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);

                wordApp.Quit(MSWord.WdSaveOptions.wdDoNotSaveChanges, MSWord.WdOriginalFormat.wdOriginalDocumentFormat, false);
            }
        }

        public void killMSWordProcess()
        {
            try
            {
                System.Diagnostics.Process.Start("taskkill", "/F /IM WINWORD.EXE");
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
            }
        }

        public void createDirectory(string htmlFolderPath)
        {
            if (!Directory.Exists(htmlFolderPath))
            {
                Directory.CreateDirectory(htmlFolderPath);
            }
        }

        //test pass
        public string html2doc(string htmlPath)
        {
            string docSavePath = "";
            try
            {
                string docSaveFolderPath = Path.GetDirectoryName(htmlPath) + "\\";
                MSWord._Application wordApp = new MSWord.Application() { Visible = false, DisplayAlerts = MSWord.WdAlertLevel.wdAlertsNone };

                //name: html->doc
                docSavePath = docSaveFolderPath + Path.GetFileNameWithoutExtension(htmlPath) + ".doc";
                if (!File.Exists(docSavePath))
                {
                    MSWord._Document doc = wordApp.Documents.Open(htmlPath);
                    //
                    wordApp = (MSWord.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
                    MSWord._Document wdDoc = wordApp.ActiveDocument;

                    //
                    doc.SaveAs2(docSavePath, MSWord.WdSaveFormat.wdFormatDocument);
                    doc.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                }

                wordApp.Quit();
            }
            catch (Exception e)
            {
                //e.Message;
            }

            return docSavePath;
        }

        //test pass
        public string doc2pdf(string docPath)
        {
            string pdfSavePath = "";
            try
            {
                string docSaveFolderPath = Path.GetDirectoryName(docPath) + "\\";
                MSWord._Application wordApp = new MSWord.Application() { Visible = false, DisplayAlerts = MSWord.WdAlertLevel.wdAlertsNone };

                //name: html->doc
                pdfSavePath = docSaveFolderPath + Path.GetFileNameWithoutExtension(docPath) + ".pdf";
                if (!File.Exists(pdfSavePath))
                {
                    MSWord._Document doc = wordApp.Documents.OpenNoRepairDialog(docPath);
                    //
                    wordApp = (MSWord.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
                    MSWord.Document wdDoc = wordApp.ActiveDocument;
                    //
                    doc.SaveAs2(pdfSavePath, MSWord.WdSaveFormat.wdFormatPDF);
                    doc.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                }

                wordApp.Quit();
            }
            catch (Exception e)
            {
                //e.Message;
            }

            return pdfSavePath;
        }

        //Note 1.
        //word Show Repairs dialog box content:
        //Show Repairs
        //Errors were detected in this file, but word was able to open the file by
        //making the repaird listed below. save the file to make the repairs permanent
        //SOLUTION: use Documents.OpenNoRepairDialog instead of Documents.Open
        //
        //Note 2.
        //warning content:
        //Ambiguity between method 'Microsoft.Office.Interop.Word._Document.Close(ref object, ref object, ref object)' 
        //and non-method 'Microsoft.Office.Interop.Word.DocumentEvents2_Event.Close'. Using method group.
        //Ambiguity between method 'Microsoft.Office.Interop.Word._Application.Quit(ref object, ref object, ref object)' 
        //and non-method 'Microsoft.Office.Interop.Word.ApplicationEvents4_Event.Quit'. Using method group.
        //SOLUTION:
        //Application and Document are subclasses that add new members to _Application and _Document.
        //In VS2010, the best way to get rid of this pesky warning is is to declare 
        //the variables _Document rather than Document and _Application rather than Application.
        //Of course, you still have to use non '_' notation when instantiating:
        //_Application wordApp = new Application();
        //_Document doc = wordApp.Documents.Open(htmlPath);
    }
}
