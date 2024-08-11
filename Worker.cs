using System.ComponentModel.DataAnnotations;
using System.Reflection;
using TimeIngest;



namespace TimeIngest;

public class Worker : BackgroundService
{
    string parentPath = "start";
    private readonly ILogger<Worker> _logger;
  
    public Worker(ILogger<Worker> logger)
    {
        _logger = logger;
                
        if (parentPath == "start") {

            string? currentPath = Path.GetDirectoryName(System.AppContext.BaseDirectory);
            if (currentPath != null) {
                parentPath = Path.GetFullPath(Path.Combine(currentPath, "."));
            }


        }

    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {

       

        while (!stoppingToken.IsCancellationRequested)
        {
            if (_logger.IsEnabled(LogLevel.Information))
            {
                _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
                _logger.LogInformation("Location: {loc}", parentPath);
            }
            await Task.Delay(1000, stoppingToken);
            FileSystemWatcher watcher = new FileSystemWatcher();
            if(parentPath != "start") {
                watcher.Path = parentPath;
            }
                
           
            watcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName;
            watcher.Filter = "*.*";

            //watcher.Changed += new FileSystemEventHandler(onChanged);
            watcher.Created += new FileSystemEventHandler(onChanged);
            // watcher.Deleted += new FileSystemEventHandler(onChanged);
            

            watcher.EnableRaisingEvents = true;


        }
    }


    public void onChanged(object source, FileSystemEventArgs e) {
        if (e is not null)
        {
                    
            _logger.LogInformation("Location: {loc}", parentPath);
            //  string? msg = e.Name != null ? ConvertMsgToJson.Convert(e.Name) : "";
            
            if(e.Name is not null && e.Name.Contains(".msg") && !e.Name.Contains("_processed_")) {
                string? dirname = Path.GetDirectoryName(e.FullPath); 
                string newfilename = dirname + "\\" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + "_processed_" + e.Name;
                try {
                    System.IO.File.Move(e.FullPath, newfilename);
                }
                catch {};
                


            }
        }
    }
}
 