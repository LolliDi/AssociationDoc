using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AssociationDoc
{
    struct FileSource
    {
        public string Path { get; set; }
        string fileName;

        public string FileName { get => fileName; set => fileName = value; }

    }
}
