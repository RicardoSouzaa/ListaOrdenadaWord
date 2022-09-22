using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;

namespace ListaOrdenadaWord
{
    public class Program
    {
        public Application wordApp = new Application();

        public Document aDoc = null;

        Paragraph paragraph = null;

        ListFormat listFormat = null;

        object readOnly = false;

        object missing = Missing.Value;

        static void Main(string[] args)
        {
            var mc = new Program();

            List<int> processesbeforegen = mc.getRunningProcesses();

            mc.CreatWordDocument("<TELEFONE>", "(43) 9991-9191");

            mc.aDoc.SaveAs(
                @"D:\Dev\dados\teste2.docx",
                ref mc.missing, ref mc.readOnly, ref mc.missing, ref mc.missing,
                ref mc.missing, ref mc.missing, ref mc.missing, ref mc.missing,
                ref mc.missing, ref mc.missing, ref mc.missing, ref mc.missing,
                ref mc.missing, ref mc.missing, ref mc.missing
            );

            mc.aDoc.Close(ref mc.missing, ref mc.missing, ref mc.missing);

            List<int> processesaftergen = mc.getRunningProcesses();
            mc.killProcesses(processesbeforegen, processesaftergen);

            Console.WriteLine("FINALIZADO --- Arquivo Criado");
            Console.ReadLine();
        }

        private void CreatWordDocument(object findText, object replaceText)
        {
            //ListGallery listGallery = wordApp.ListGalleries[WdListGalleryType.wdOutlineNumberGallery];

            wordApp.Visible = false;

            aDoc = wordApp.Documents.Open(
                @"D:\Dev\dados\teste.docx",
                ref missing, ref readOnly, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, false, ref missing,
                ref missing, ref missing, ref missing
            );

            //Range range = aDoc.Content;
            Range range = aDoc.Range();

            string sList = "";
            List oLst = aDoc.Lists[1];

            for (int i = 1; i <= oLst.ListParagraphs.Count; i++)
            {
                sList += oLst.ListParagraphs[i].Range.Text + "\n";
            }
            Console.WriteLine(sList);

            ///// ############################################################################## FUNCIONAAAA lista com multilevel substituindo uma tag.

            string search = "$list";
            while (range.Find.Execute(search))
            {
                ListGallery listGallery =
                    wordApp.ListGalleries[WdListGalleryType.wdNumberGallery];

                // Select found location
                range.Select();

                // Apply multi level list
                wordApp.Selection.Range.ListFormat.ApplyListTemplateWithLevel
                (
                    listGallery.ListTemplates[1],
                    ContinuePreviousList: true,
                    ApplyTo: WdListApplyTo.wdListApplyToWholeList,
                    DefaultListBehavior: WdDefaultListBehavior.wdWord10ListBehavior
                 );

                // First level
                wordApp.Selection.TypeText("Root Item A".ToUpper());  // Set text to key in
                wordApp.Selection.TypeParagraph();  // Simulate typing in MS Word

                // Go to 2nd level
                wordApp.Selection.Range.ListFormat.ListIndent();
                wordApp.Selection.TypeText("Child Item A.1");
                wordApp.Selection.
                wordApp.Selection.TypeParagraph();
                wordApp.Selection.TypeText("Child Item A.2");
                wordApp.Selection.TypeParagraph();

                // Back to 1st level
                wordApp.Selection.Range.ListFormat.ListOutdent();
                wordApp.Selection.TypeText("Root Item B".ToUpper());
                wordApp.Selection.TypeParagraph();

                //Go to 2nd level
                wordApp.Selection.Range.ListFormat.ListIndent();
                wordApp.Selection.TypeText("Child Item B.1");
                wordApp.Selection.TypeParagraph();
                wordApp.Selection.TypeText("Child Item B.2");
                wordApp.Selection.TypeParagraph();

                wordApp.Selection.TypeBackspace();
            }
            ///// ############################################################################## FUNCIONAAAA
            //string testRows = "Test 1\n\tTest 2\tName\tAmount\nTest 3\nTest 4\nTest 5\nTest 6\nTest 7\nTest 8\nTest 9\nTest 10\n";

            //try
            //{
            //    //adiciona um paragrafo depois do 1. e antes do 2.
            //    paragraph = aDoc.ListParagraphs[1].Next();
            //    paragraph = aDoc.Paragraphs.Add(paragraph.Range);
            //    paragraph.Range.Text = testRows;
            //    paragraph.Range.InsertParagraphAfter();

            //    paragraph = aDoc.ListParagraphs[11].Next();
            //    paragraph = aDoc.Paragraphs.Add(paragraph.Range);
            //    paragraph.Range.Text = "Teste 11\n";
            //    paragraph.Range.ListFormat.ApplyOutlineNumberDefault(WdDefaultListBehavior.wdWord10ListBehavior);
            //    paragraph.Outdent();
            //    paragraph.Range.InsertParagraphAfter();

            //    paragraph = aDoc.Paragraphs.Add(paragraph.Range);
            //    paragraph.Range.Text = "Teste 11.1";
            //    paragraph.Indent();
            //    paragraph.Range.InsertParagraphAfter();
            //}
            //catch (System.Runtime.InteropServices.COMException e)
            //{
            //    Console.WriteLine("COMException: " + e.StackTrace.ToString());
            //    Console.ReadKey();
            //}

            ///// ##############################################################################

            //aDoc.Range().ListFormat.ApplyListTemplateWithLevel
            //(
            //    ListTemplate: aDoc.ListTemplates[listNumber],
            //    ContinuePreviousList: true,
            //    ApplyTo: WdListApplyTo.wdListApplyToSelection,
            //    DefaultListBehavior: WdDefaultListBehavior.wdWord10ListBehavior
            //);

            //////############################################################################# modelo para aplicar o template por método

            //paragraph = range.Paragraphs.Add();
            //listFormat = paragraph.Range.ListFormat;
            //paragraph.Range.Text = "Root Item A";
            //this.ApplyListTemplate(listGallery, listFormat, 1);
            //paragraph.Range.InsertParagraphAfter();

            //paragraph = paragraph.Range.Paragraphs.Add();
            //listFormat = paragraph.Range.ListFormat;
            //paragraph.Range.Text = "Child Item A.1";
            //this.ApplyListTemplate(listGallery, listFormat, 2);
            //paragraph.Range.InsertParagraphAfter();
            //////#############################################################################
        }

        private void ApplyListTemplate(ListGallery listGallery, ListFormat listFormat, int level = 1)
        {
            listFormat.ApplyListTemplateWithLevel(
                listGallery.ListTemplates[level],
                ContinuePreviousList: true,
                ApplyTo: WdListApplyTo.wdListApplyToSelection,
                DefaultListBehavior: WdDefaultListBehavior.wdWord10ListBehavior,
                ApplyLevel: level);
        }

        public List<int> getRunningProcesses()
        {
            List<int> ProcessIDs = new List<int>();

            foreach (Process clsProcess in Process.GetProcesses())
            {
                if (Process.GetCurrentProcess().Id == clsProcess.Id)
                    continue;
                if (clsProcess.ProcessName.Contains("WINWORD"))
                {
                    ProcessIDs.Add(clsProcess.Id);
                }
            }
            return ProcessIDs;
        }

        private void killProcesses(List<int> processesbeforegen, List<int> processesaftergen)
        {
            foreach (int pidafter in processesaftergen)
            {
                bool processfound = false;
                foreach (int pidbefore in processesbeforegen)
                {
                    if (pidafter == pidbefore)
                    {
                        processfound = true;
                    }
                }

                if (processfound == false)
                {
                    Process clsProcess = Process.GetProcessById(pidafter);
                    clsProcess.Kill();
                }
            }
        }
    }
}