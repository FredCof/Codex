// using System.IO.Packaging;

using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Codex.Word.Net.Utils;

namespace Codex.Word.Net.Base
{
    internal class NodeTree
    {
        private class Node
        {
            private int _hashId;
            private string _name;
            private List<Node> _children;

            public int HashIndex {
                get => _hashId;
                set => _hashId=value;
            }

            public string Name
            {
                get => _name;
                set => _name = value;
            }

            public Node(string name)
            {
                _name = name;
                _children = new List<Node>();
                // ReSharper disable once CA1307
                if (name != null) _hashId = name.GetHashCode();
            }

            public void Append(Node node)
            {
                _children.Append(node);
            }

            public void Prepend(Node node)
            {
                _children.Prepend(node);
            }

            public void Clear()
            {
                _children.Clear();
            }
        }
    }

    /// <summary>
    /// All office elements base on BaseElement 
    /// </summary>
    public abstract class BaseElement
    {
        #region Members

        private XElement Xml { get; set; }

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
    }
}