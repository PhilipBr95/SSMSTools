﻿
using System.Collections.Generic;

namespace SsmsSchemaFolders
{
    public interface ISchemaFolderOptions
    {
        bool Enabled { get; }
        bool AppendDot { get; set; }
        bool CloneParentNode { get; }
        bool UseObjectIcon { get; }
        List<string> Separators { get; }
    }
}
