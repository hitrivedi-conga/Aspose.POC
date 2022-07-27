using Conga.DocGen.DataSet;
using Conga.DocGen.DataSet.Models;

namespace Aspose.POC.XATemplateExample {
    /// <summary>
    /// Represents the data source that is used for merging with template.
    /// </summary>
    public class TemplateDataSource : ITemplateDataSource {
        /// <summary>
        /// Smart DataSet.
        /// </summary>
		public SmartSet DataSet { get; }

        /// <summary>
        /// Constructor.
        /// </summary>
        public TemplateDataSource() {
            DataSet = new SmartSet();
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="dataSet">DataSet</param>
        public TemplateDataSource(SmartSet dataSet) {
            DataSet = dataSet;
        }

        /// <summary>
        /// Fills the smart set with schema and data.
        /// </summary>
        /// <param name="schemaAndDataXml">Schema and data xml.</param>
        public void FillSmartSet(string schemaAndDataXml) {
            DataSet.Fill(schemaAndDataXml);
        }

        /// <summary>
        /// Retrieved data from the passed path and then formats the retrieve data based on the passed format
        /// or from the default value configured for the data type.
        /// </summary>
        /// <param name="path">Data path</param>
        /// <returns>Formatted data</returns>
        public string Render(string path) {
            return DataSet.Render(path);
        }

        /// <summary>
        /// Returns the binded entity or field metadata from the passed path.
        /// </summary>
        /// <typeparam name="T">Metadata type</typeparam>
        /// <param name="path">XPath</param>
        /// <returns>Entity Metadata or Field Metadata</returns>
        public T GetMetadata<T>(string path) where T : Metadata {
            return DataSet.GetMetadata<T>(path);
        }
    }
}
