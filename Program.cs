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

        static void Main(string[] args)
        {
            var mc = new Program();

            Console.WriteLine("FINALIZADO --- Arquivo Criado");
            Console.ReadLine();
        }
    }
}