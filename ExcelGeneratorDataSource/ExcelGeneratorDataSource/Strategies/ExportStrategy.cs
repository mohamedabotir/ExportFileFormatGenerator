using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelGeneratorDataSource.Strategies
{
    public interface ExportStrategy
    {

        public string Export(object datasource);

    }
}
