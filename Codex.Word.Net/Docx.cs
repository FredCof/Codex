/*
 * Codex – Codex is an open source word operation lib for .NET
 * Copyright (C) 2020-2020 Cofalconer.
 *
 * author: Cofalconer
 * date: 2020-03-25
 * summary: Word Operation
 */
using System;

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
            Console.WriteLine(ofDocx);
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
