// using System.IO.Packaging;

using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Resources;
using System.Text;
using System.Xml.Linq;
using Codex.Word.Net.Utils;

namespace Codex.Word.Net.Base
{
    internal class XML
    {
        public XElement Xml { get; set; }

        public XML()
        {
            return;
        }
    }

    /// <summary>
    /// All office elements base on BaseElement 
    /// </summary>
    public abstract class BaseElement
    {
        #region Members

        /// <summary>
        /// This is the XML substance, with will contain:w:p>w:r>w:t
        /// </summary>
        public XElement Xml { get; set; }

        #endregion

        #region Constructors

        public BaseElement()
        {
            // throw new NotImplementedException(Resource.NoOverride);
            return;
        }

        #endregion

        #region Normal Method

        /// <summary>
        /// Generate the various ids in the document, such as rsid, textid
        /// </summary>
        /// <returns>An eight-character string consisting of uppercase letters and Numbers</returns>
        private string GetRandomId()
        {
            Random ran=new Random();
            StringBuilder randomId = new StringBuilder();
            for (int i = 0; i < 8; i++)
            {
                int ascii = ran.Next(1, 36) > 10 ? ran.Next(65, 90) : ran.Next(48, 57);
                randomId.Append((char)ascii);
            }
            return randomId.ToString();
        }

        #endregion

        #region Virtual Methods

        public virtual BaseElement Query(string selector)
        {
            throw new NotImplementedException(Resource.NoOverride);
        }

        #endregion
    }
}