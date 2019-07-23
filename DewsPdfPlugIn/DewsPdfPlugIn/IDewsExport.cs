using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DewsPdfPlugIn
{
    interface IDewsExport
    {
        void Export(IDictionary<string, string> ProjectDetails, IDictionary<string, string> Metrics, IDictionary<string, Dictionary<string, string>> ProjectMetricValues, string filepath);
    }
}
