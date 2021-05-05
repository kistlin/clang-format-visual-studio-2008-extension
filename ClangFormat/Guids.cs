using System;

namespace LLVM.ClangFormat
{
    static class GuidList
    {
        public const string guidClangFormatPkgString = "bb9ab141-fda4-4c6b-b243-ff097b1c2f54";
        public const string guidClangFormatCmdSetString = "902a7566-cb15-4838-a088-75b6ee238fff";

        public static readonly Guid guidClangFormatCmdSet = new Guid(guidClangFormatCmdSetString);
    };
}