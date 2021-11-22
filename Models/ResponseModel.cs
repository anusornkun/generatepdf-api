namespace iTextGenPDF.Api.Models
{
    public class ResponseModel
    {
        public int StatusCode { get; set; }
        public object Result { get; set; }
    }

    public class FileResponseDataBinding
    {
        public int StatusCode { get; set; }
        public string Message { get; set; }
        public byte[] File { get; set; }
    }
}
