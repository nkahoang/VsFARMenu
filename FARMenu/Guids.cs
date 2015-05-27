// Guids.cs
// MUST match guids.h
using System;

namespace nkaHnt.FARMenu
{
    static class GuidList
    {
        public const string guidFARMenuPkgString = "85caa711-8e6b-45a7-bc6c-afbc56605309";
        public const string guidFARMenuCmdSetString = "cbb095bb-f1c7-4da2-b11c-460ae5295cf8";

        public static readonly Guid guidFARMenuCmdSet = new Guid(guidFARMenuCmdSetString);
    };
}