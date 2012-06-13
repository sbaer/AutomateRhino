using System;

namespace AutomateRhino_CS
{
  class Program
  {
    static void Main(string[] args)
    {
      Console.Write("Starting Rhino ");
      System.Type t = System.Type.GetTypeFromProgID("Rhino5x64.Application");
      dynamic rh = System.Activator.CreateInstance(t);
      rh.ReleaseWithoutClosing = 0;

      // wait until Rhino is done initializing
      // bail out after 15 seconds if Rhino doesn't initialize in that period of time
      const int bail_milliseconds = 15 * 1000;
      int time_waiting = 0;
      while (rh.IsInitialized() == 0)
      {
        Console.Write(".");
        System.Threading.Thread.Sleep(100);
        time_waiting += 100;
        if (time_waiting > bail_milliseconds)
          return;
      }

      Console.WriteLine();
      //rh.Visible = 1; //uncomment if you want to see Rhino's UI

      // If you want to open a specific 3dm file, you could run the following
      //
      //string input_file = @"C:\Users\a-steve\Desktop\3dm models\curve weight.3dm";
      //Console.WriteLine("Reading " + input_file);
      //rh.RunScript("-Open \"" + input_file + "\"", 0);
      

      // Get a hold of the grasshopper COM object
      const string gh_plugin_id = "b45a29b1-4343-4035-989e-044e8580d9cf";
      dynamic grasshopper = rh.GetPlugInObject(gh_plugin_id, Guid.Empty.ToString());

      const string gh_doc = @"C:\Users\a-steve\Desktop\simpledef.gh";
      grasshopper.OpenDocument(gh_doc);

      rh.Visible = 1; //uncomment if you want to see Rhino's UI

      // If you want to save a specific 3dm file, you could run the following
      //string output_file = @"C:\Users\a-steve\Desktop\3dm models\curve weight.dxf";
      //Console.WriteLine("Writing " + output_file);
      //rh.RunScript("-Save \"" + output_file + "\" Enter", 0);

      Console.WriteLine("Done");
    }
  }
}
