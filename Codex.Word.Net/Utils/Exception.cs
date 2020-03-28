using System;
using System.Runtime.Serialization;

namespace Codex.Word.Net.Utils
{
    [Serializable]
    public class BaseDocxException : Exception
    {
        //
        // For guidelines regarding the creation of new exception types, see
        //    http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cpgenref/html/cpconerrorraisinghandlingguidelines.asp
        // and
        //    http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dncscol/html/csharp07192001.asp
        //

        public BaseDocxException()
        {
        }

        public BaseDocxException(string message) : base(message)
        {
        }

        public BaseDocxException(string message, Exception inner) : base(message, inner)
        {
        }

        protected BaseDocxException(
            SerializationInfo info,
            StreamingContext context) : base(info, context)
        {
        }
    }
}