using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using CrystalDecisions.CrystalReports.Engine;
using System.IO;
using System.Collections;
using CrystalDecisions.Shared;

namespace z.IO.Report
{
    /// <summary>
    /// LJ 20150824
    /// </summary>
    public class Crystal : IDisposable
    {
        /// <summary>
        /// DataSource
        /// </summary>
        public DataTable DataSource { private get; set; }
        public string ReportFile { private get; set; }
        private Dictionary<string, object> rParameters;
        public string ReportData { private set; get; }

        public ReportType Type { private get; set; } = ReportType.PortableDocFormat;

        public delegate void Rpt(ReportDocument rpt); //nat 20160630

        public List<SubReportSourceCtx> SubReportSource { get; set; } = new List<SubReportSourceCtx>();

        /// <summary>
        /// Creates a shadow copy of report and use as template
        /// </summary>
        //public bool MultiTasking { private get; set; } = false;

        public string ReportName
        {
            get
            {
                return Path.GetFileName(ReportFile);
            }
        }

        public DBParameters Credentials { private get; set; }

        public ReportDocument rptDoc { get; private set; }
        private string tmp;

        public Crystal()
        {
            this.rParameters = new Dictionary<string, object>();
            this.rptDoc = new ReportDocument();
        }

        /// <summary>
        /// Prepare the report document
        /// </summary>
        /// <param name="MultiTask">
        /// Creates a shadow copy of report and use as template
        /// </param>
        public void Init(bool MultiTask = true)
        {
            if (File.Exists(this.ReportFile))
            {
                if (MultiTask)
                {
                    tmp = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".rpt");
                    File.Copy(this.ReportFile, tmp);
                    rptDoc.Load(tmp);
                }
                else
                    rptDoc.Load(this.ReportFile);
            }
            else throw new Exception("Can not find the Report File");
        }

        //public void Init(Stream rptFile)
        //{
        //    rptDoc.Load( //.Load(rptFile);

        //}

        /// <summary>
        /// Load the Report and Generate a Base64 Data
        /// </summary>
        /// <param name="report"></param>
        public void Load(Rpt report = null)
        {
            try
            {
                if (DataSource == null) throw new Exception("Please provide Report Data Source");

                report?.Invoke(rptDoc); //nat 20160630
                rptDoc.SetDatabaseLogon(Credentials.User, Credentials.Password, Credentials.Server, Credentials.Database);
                rptDoc.SetDataSource(this.DataSource);

                foreach (ReportDocument ireport in rptDoc.Subreports)
                    foreach (IConnectionInfo dsc in ireport.DataSourceConnections)
                        dsc.SetConnection(Credentials.Server, Credentials.Database, Credentials.User, Credentials.Password);

                if (SubReportSource.Count > 0)
                {
                    SubReportSource.ForEach(x =>
                    {
                        var gg = rptDoc.Subreports.Cast<ReportDocument>().Where(y => y.Name == x.Name);
                        if (gg.Any())
                            gg.Single().SetDataSource(x.DataSource);
                    });
                }

                //if(SubReportSource.Count > 0)
                //{   
                //    foreach(ReportDocument rd in rptDoc.Subreports)
                //    { 
                //        if (SubReportSource.Any(x => x.Name == rd.Name))
                //            rd.SetDataSource(SubReportSource.Single(x => x.Name == rd.Name));
                //        else
                //            foreach (IConnectionInfo dsc in rd.DataSourceConnections)
                //                dsc.SetConnection(Credentials.Server, Credentials.Database, Credentials.User, Credentials.Password);
                //    }
                //}


                var ggs = GetParameterKeys();
                foreach (var g in this.rParameters)
                    if (ggs.Contains(g.Key))
                        rptDoc.SetParameterValue(g.Key, g.Value);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (File.Exists(tmp))
                    File.Delete(tmp);
                GC.Collect();
            }
        }

        public void Save()
        {
            using (var ms = new MemoryStream())
            {
                var format = string.Empty;
                switch (Type)
                {
                    case ReportType.PortableDocFormat:
                        rptDoc.ExportToStream(ExportFormatType.PortableDocFormat).CopyTo(ms);
                        format = "application/pdf";
                        break;
                    case ReportType.Excel:
                        rptDoc.ExportToStream(ExportFormatType.Excel).CopyTo(ms);
                        format = "application/vnd.ms-excel";
                        break;
                }
                ms.Position = 0;
                byte[] buffer = ms.ToArray();
                ms.Close();
                ReportData = $"data:{format};base64,{Convert.ToBase64String(buffer)}";
            }
        }

        public void Save(Stream ms)
        {
            var format = string.Empty;
            switch (Type)
            {
                case ReportType.PortableDocFormat:
                    rptDoc.ExportToStream(ExportFormatType.PortableDocFormat).CopyTo(ms);
                    format = "application/pdf";
                    break;
                case ReportType.Excel:
                    rptDoc.ExportToStream(ExportFormatType.Excel).CopyTo(ms);
                    format = "application/vnd.ms-excel";
                    break;
            }
        }

        public void AddParameters(string key, object value)
        {
            rParameters.Add(key, value);
        }

        public void AddParameters(params KeyValuePair<string, object>[] args)
        {
            foreach (var h in args)
                this.AddParameters(h.Key, h.Value);
        }

        public IEnumerable<string> GetParameterKeys()
        {
            return rptDoc.ParameterFields.Cast<ParameterField>().Select(x => x.Name);
        }

        public IEnumerable<ParameterCtx> GetPromptKeys()
        {
            return
                rptDoc.ParameterFields.Cast<ParameterField>()
                    .Where(x => (x.ParameterFieldUsage2 & (ParameterFieldUsage2.ShowOnPanel)) != 0) //ParameterFieldUsage2.InUse |
                    .Select(x => new ParameterCtx()
                    {
                        Name = x.Name,
                        Label = x.PromptText
                    });
        }

        public List<string> GetSubReportTableNames()
        {
            return rptDoc.Subreports.Cast<ReportDocument>().Select(x => x.Name).ToList();
        }

        public void Dispose()
        {
            DataSource?.Dispose();
            rParameters = null;
            this.rptDoc?.Dispose();
            ReportData = "";

            GC.Collect();
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Report Type
        /// </summary>
        public enum ReportType
        {
            NoFormat = 0,
            CrystalReport = 1,
            RichText = 2,
            WordForWindows = 3,
            Excel = 4,
            PortableDocFormat = 5,
            HTML32 = 6,
            HTML40 = 7,
            ExcelRecord = 8,
            Text = 9,
            CharacterSeparatedValues = 10,
            TabSeperatedText = 11,
            EditableRTF = 12,
            Xml = 13,
            RPTR = 14,
            ExcelWorkbook = 15
        }
    }
}
