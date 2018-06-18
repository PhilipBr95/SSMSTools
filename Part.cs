using System.Diagnostics;

namespace SsmsSchemaFolders
{
    [DebuggerDisplay("{FullName}")]
    class Part
    {
        public Part Sibling { get; set; }
        public string FullName { get; set; }
        public string Name { get; set; }
        public string Separator { get; set; }
        public int Location { get; set; }

        internal Part GetLastPart()
        {
            if (Sibling == null)
                return this;

            return this.Sibling.GetLastPart();
        }
    }
}
