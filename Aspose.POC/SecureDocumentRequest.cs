namespace Aspose.POC
{
    /// <summary>
    /// This class contains all properties that are required to apply or remove security from or to a document.
    /// </summary>
    public class SecureDocumentRequest
    {
        /// <summary>
        /// The document id to which we need to apply or remove protection.
        /// </summary>
        public string DocumentId { get; set; }

        /// <summary>
        /// The password to apply to protect the document.
        /// </summary>
        public string Password { get; set; }

        public bool AllowPrinting { get; set; }
        public bool AllowCopying { get; set; }
        public bool AllowCommenting { get; set; }
        public bool AllowPageExtraction { get; set; }
        public bool AllowFormFilling { get; set; }
        public bool AllowSigning { get; set; }
        public string DocFormat { get; set; }

        public bool SecurityEnabled { get; set; }

    }


}