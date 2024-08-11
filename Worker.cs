using System.ComponentModel.DataAnnotations;
using System.Reflection;
using TimeIngest;
using Python.Runtime;



namespace TimeIngest;

public class Worker : BackgroundService
{
    string parentPath = "start";
    private readonly ILogger<Worker> _logger;
  
    public Worker(ILogger<Worker> logger)
    {
        _logger = logger;
        
         parentPath = Helper.GetExecutionPath();       
        

    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {

       

        while (!stoppingToken.IsCancellationRequested)
        {
            if (_logger.IsEnabled(LogLevel.Information))
            {
               // _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
               // _logger.LogInformation("Location: {loc}", parentPath);
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
                    
           
            
        
            
            
                
            

           
               
                


            
        }
    }
}
 