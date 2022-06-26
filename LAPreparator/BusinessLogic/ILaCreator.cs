using LAPreparator.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace LAPreparator.BusinessLogic
{
    public interface ILaCreator
    {
        List<LA> CreateLAs(
            IEnumerable<IGrouping<string, DataRow>> groupedData,
            string sourceFile,string sourceSheet,string sourceRange
            //string destinationFile,string destinationSheet,string destinationRange
            );
        
    }
}
