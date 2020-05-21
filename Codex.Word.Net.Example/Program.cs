using System;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using Codex.Word.Net.Base;

namespace Codex.Word.Net.Example
{
    class Program
    {
        static void Main(string[] args)
        {
            Docx x = new Docx();
            string style = Base.BaseStyle.Normal.GetDescription();
            Console.WriteLine(style);
        }
    }
}
