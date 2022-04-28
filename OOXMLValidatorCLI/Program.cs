using DocumentFormat.OpenXml;
using Microsoft.Extensions.DependencyInjection;
using OOXMLValidatorCLI.Classes;
using OOXMLValidatorCLI.Interfaces;
using System;

namespace OOXMLValidatorCLI
{
    class Program
    {
        private static void Main(string[] args)
        {
            // set up DI
            ServiceCollection collection = new ServiceCollection();
            collection.AddScoped<IValidate, Validate>();
            collection.AddScoped<IFunctionUtils, FunctionUtils>();
            collection.AddScoped<IDocumentUtils, DocumentUtils>();
            collection.AddSingleton<IFileService, FileService>();
            collection.AddSingleton<IDirectoryService, DirectoryService>();

            ServiceProvider serviceProvider = collection.BuildServiceProvider();

            IValidate validate = serviceProvider.GetService<IValidate>();

            string xmlPath;
            bool returnXml = false;
            string version = null;
            bool recursive = false;

            if (args != null && args.Length > 0)
            {
                xmlPath = args[0];

                for (int i = 1; i < args.Length; i++)
                {
                    if (Enum.TryParse(args[i], out FileFormatVersions v))
                    {
                        version = args[i];
                    }
                    else
                    {
                        switch (args[i])
                        {
                            case "--xml":
                                returnXml = true;
                                break;
                            case "-r":
                                recursive = true;
                                break;
                            case "--recursive":
                                recursive = true;
                                break;
                            default: throw new ArgumentException("Unknown argument", args[i]);
                        }
                    }
                }
            }
            else
            {
                throw new ArgumentNullException();
            }


            if (args.Length > 2)
            {
                returnXml = args[1] == "--xml" || args[2] == "--xml" ? true : false;
            }

            object validationErrors = validate.OOXML(xmlPath, version, returnXml, recursive);

            Console.Write(validationErrors);
        }
    }
}
