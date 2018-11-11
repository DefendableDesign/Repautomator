namespace Repautomator
{
    public class ServerResponse
    {
        public string Content { get; set; }
        public int Status { get; set; }

        /// <summary>
        /// ServerResponse Constructor
        /// </summary>
        /// <param name="content">Raw content/result from a request.</param>
        /// <param name="status">HTTP Response status code.</param>
        public ServerResponse(string content, int status)
        {
            Content = content;
            Status = status;
        }
    }
}
