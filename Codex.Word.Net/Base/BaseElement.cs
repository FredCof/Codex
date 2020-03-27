// using System.IO.Packaging;
using System.IO.Packaging;
using System.Xml.Linq;

namespace Codex.Word.Net.Base
{
    /// <summary>
    /// All office elements base on BaseElement 
    /// </summary>
    public abstract class BaseElement
    {
        #region Members

        private XElement Xml { get; set; }

        #endregion
    }
}