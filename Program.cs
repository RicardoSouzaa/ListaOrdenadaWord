using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;

using Word = Microsoft.Office.Interop.Word;

namespace ListaOrdenadaWord
{
    public class Program
    {
        public Application wordApp = new Application();
        public Document aDoc = null;

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
            //wordApp.Documents.Add(ref missing, ref missing,
            //ref missing, ref missing);

            //wordApp.Visible = true;

            wordApp.Visible = false;

            //aDoc = wordApp.Documents.Open(
            //    @"D:\Dev\dados\teste.docx",
            //    ref missing, ref readOnly, ref missing, ref missing,
            //    ref missing, ref missing, ref missing, ref missing,
            //    ref missing, ref missing, false, ref missing,
            //    ref missing, ref missing, ref missing
            //);

            //const int listNumber = 1; //The first list on the page is list 1, the second is list 2 etc etc

            //aDoc = wordApp.Documents.Open(@"D:\Dev\dados\teste.docx");
            aDoc = wordApp.Documents.Open(
                @"D:\Dev\dados\teste.docx",
                ref missing, ref readOnly, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, false, ref missing,
                ref missing, ref missing, ref missing
            );

            Range range = aDoc.Content;

            string sList = "";
            List oLst = aDoc.Lists[1];
            //Range startList = oLst.Range;
            //Range endList = startList.Duplicate;
            //startList.End = endList.End;

            for (int i = 1; i <= oLst.ListParagraphs.Count; i++)
            {
                sList += oLst.ListParagraphs[i].Range.Text + "\n";
            }
            Console.WriteLine(sList);

            //aqui define onde vai começar a lista.
            range = aDoc.ListParagraphs[1].Next().Range;
            range.InsertParagraphAfter();
            //range.ListParagraphs[1].Range.InsertParagraphAfter();

            int startOfList = range.Start;

            range.Text = "teste\nteste 2\nteste 3";
            range.InsertParagraphAfter();
            range.InsertParagraphAfter();
            int endOfList = range.End;

            Range listRange = aDoc.Range(startOfList, endOfList);
            listRange.ListFormat.ApplyNumberDefault();

            //range = aDoc.Paragraphs.Last.Range;
            //range.Text = "Bye for now!";
            //range.InsertParagraphAfter();

            //////#######################################################################
            //aDoc.Range().ListFormat.ApplyListTemplateWithLevel
            //(
            //    ListTemplate: aDoc.ListTemplates[listNumber],
            //    ContinuePreviousList: true,
            //    ApplyTo: WdListApplyTo.wdListApplyToSelection,
            //    DefaultListBehavior: WdDefaultListBehavior.wdWord10ListBehavior
            //);

            //////#############################################################################
            //Document doc = wordApp.Documents.Add();

            //Range range = doc.Content;
            //range.Text = "Hello world!";

            //range.InsertParagraphAfter();
            //range = doc.Paragraphs.Last.Range;

            //// start of list
            //int startOfList = range.Start;

            //// each \n character adds a new paragraph...
            //range.Text = "Item 1\nItem 2\nItem 3";

            //// ...or insert a new paragraph...
            //range.InsertParagraphAfter();
            //range = doc.Paragraphs.Last.Range;
            //range.Text = "Item 4\nItem 5";

            //// end of list
            //int endOfList = range.End;

            //// insert the next paragraph before applying the format, otherwise
            //// the format will be copied to the suceeding paragraphs.
            //range.InsertParagraphAfter();

            //// apply list format
            //Range listRange = doc.Range(startOfList, endOfList);
            //listRange.ListFormat.ApplyBulletDefault();

            //range = doc.Paragraphs.Last.Range;
            //range.Text = "Bye for now!";
            //range.InsertParagraphAfter();

            //////#############################################################################
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