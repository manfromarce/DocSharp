using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocSharp.IO;

public static class ResourceDownloader
{
    public static byte[]? DownloadImage(string url)
    {
        using (var client = new System.Net.Http.HttpClient())
        {
            // Fix issue with servers refusing connections from clients without a user agent
            client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0");
            var response = client.GetAsync(url).Result;
            if (response.IsSuccessStatusCode)
            {
                var imageBytes = response.Content.ReadAsByteArrayAsync().Result;
                return imageBytes;
            }
        }
        return null;
    }

}
