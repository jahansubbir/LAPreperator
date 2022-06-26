using System;
using System.Collections.Generic;

namespace LAPreparator.Serivices
{
    public interface ITemplateReader
    {
    MessageModel Read(string fileName);
    }
}
