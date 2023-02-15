using System.Net;

public class Downloader
{
    string url = "";

    public string GetPath(string url)
    {

        string appDirectory = System.AppDomain.CurrentDomain.BaseDirectory;

        // link of the excel
        this.url = url;

        // although it is not working with link given in task
        // this method works with excel example from microsoft site
        // url = "https://go.microsoft.com/fwlink/?LinkID=521962";

        // define random extension to prevent colision in a file name
        Random random = new Random();
        int randomNumber = random.Next(1, 999);
        string fileName = @"example" + randomNumber + @".xlsx";

        // set destination in data folder
        string destination = appDirectory + @"..\..\..\data\";

        string destinationFileName = destination + fileName;

        // creaating webclient to download the files
        WebClient client = new WebClient();

        try
        {
            Console.WriteLine("Trying to download excel file from a link...");
            Console.WriteLine("Please be patient.");
            // downloading file and saving it in destination directory
            client.DownloadFile(url, destinationFileName);

            Console.WriteLine("Downloaded file: " + fileName);
        }
        catch (WebException ex)
        {
            // handling error from a link problem
            Console.WriteLine("Error downloading file: " + fileName);
            Console.WriteLine(ex.Message);

            // switch to backup
            Console.WriteLine("Backup starting...");
            destinationFileName = appDirectory + @"..\..\..\data\Worldwide Rig Count Jan 2023.xlsx";
        }

        Console.WriteLine("Finished downloading process.");

        return destinationFileName;
    }

}

