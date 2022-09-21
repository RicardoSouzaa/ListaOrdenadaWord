using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;

namespace ListaOrdenadaWord
{
    public class Program
    {
        public Word.Application wordApp = new Word.Application();
        public Word.Document aDoc = null;

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
            //List<int> processesbeforegen = getRunningProcesses();

            //object readOnly = false;

            wordApp.Visible = false;

            aDoc = wordApp.Documents.Open(
                @"D:\Dev\dados\teste.docx",
                ref missing, ref readOnly, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, false, ref missing,
                ref missing, ref missing, ref missing
            );
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