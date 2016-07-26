// Guids.cs
// MUST match guids.h
using System;

namespace Dragonist.ContextCopy
{
    static class GuidList
    {
        public const string guidContextCopyPkgString = "6539e06f-6279-42cc-99db-fadee63f38b2";
        public const string guidContextCopyCmdSetString = "fc7cb073-92a1-41bc-a2fe-d650f3de0086";

        public static readonly Guid guidContextCopyCmdSet = new Guid(guidContextCopyCmdSetString);
    };
}