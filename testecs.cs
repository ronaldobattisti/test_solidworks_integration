using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SolidWorksAutomation
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create an instance of SolidWorks application
            SldWorks swApp = null;
            try
            {
                swApp = Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")) as SldWorks;
                if (swApp == null)
                {
                    Console.WriteLine("Failed to create SolidWorks application instance.");
                    return;
                }

                // Make SolidWorks visible
                swApp.Visible = true;

                // Open a SolidWorks document (change the file path to your SolidWorks document)
                // Convert the object to ModelDoc2 explicitly
                if (swDoc != null && swDoc is ModelDoc2)
                {
                    ModelDoc2 solidWorksDocument = (ModelDoc2)swDoc;
                    // Agora você pode usar solidWorksDocument para acessar os membros e métodos de ModelDoc2
                }
                else
                {
                    Console.WriteLine("Falha ao abrir o documento do SolidWorks.");
                    return;
                }

                Console.WriteLine("SolidWorks document opened successfully.");
                
                // Your automation code here...

                // Close the SolidWorks document
                swDoc.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                // Close SolidWorks application
                if (swApp != null)
                {
                    swApp.ExitApp();
                    swApp = null;
                }
            }
        }
    }
}
