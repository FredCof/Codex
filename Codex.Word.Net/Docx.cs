/*
 * Codex – Codex is an open source word operation lib for .NET
 * Copyright (C) 2020-2020 Cofalconer.
 *
 * author: Cofalconer
 * date: 2020-03-25
 * summary: Word Operation
 */
using System;
using Codex.Word.Net.Components;
using Codex.Word.Net.Utils;

namespace Codex.Word.Net
{
    /// <summary>
    /// This is a Base Class for Docx, it will store a Docx by Stream.
    /// </summary>
    public class Docx
    {
        #region Constructor
        
        public Docx()
        {
            string ofDocx = "I am a constructor of Docx";
            // ReSharper disable once CA1303
            Console.WriteLine(ofDocx);
            Unit a = new Unit(10, SUnit.Centi);
            Unit b = Unit.Add(a, a);
            Console.WriteLine(b);
            Paragraph codexP = new Paragraph();
        }

        #endregion

        #region Methods

        #region Public Methods
        
        Boolean Init()
        {
            return false;
        }
        
        #endregion

        #region OverRide Methods
        #endregion

        #region Internal Methods
        #endregion
        
        #endregion
    }
}
