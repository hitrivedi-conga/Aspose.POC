using Conga.DocGen.DataSet.Models;

namespace Aspose.POC.XATemplateExample {
    /// <summary>
    /// Represents the data source that is used for merging with template.
    /// </summary>
	public interface ITemplateDataSource {
        /// <summary>
        /// Fills the smart set with schema and data.
        /// </summary>
        /// <param name="schemaDataXml">Schema and data xml.</param>
        void FillSmartSet(string schemaDataXml);

        /// <summary>
        /// Retrieved data from the passed path and then formats the retrieve data based on the passed format
        /// or from the default value configured for the data type.
        /// </summary>
        /// <param name="path">Data path</param>
        /// <returns>Formatted data</returns>
        string Render(string path);

        /// <summary>
        /// Returns the binded entity or field metadata from the passed path.
        /// </summary>
        /// <typeparam name="T">Metadata type</typeparam>
        /// <param name="path">XPath</param>
        /// <returns>Entity Metadata or Field Metadata</returns>
        T GetMetadata<T>(string path) where T : Metadata;
    }
}
