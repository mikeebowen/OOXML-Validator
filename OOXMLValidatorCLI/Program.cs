using DocumentFormat.OpenXml;
using Microsoft.Extensions.DependencyInjection;
using OOXMLValidatorCLI.Classes;
using OOXMLValidatorCLI.Interfaces;
using System;
using System.Threading;

namespace OOXMLValidatorCLI
{
    class Program
    {
        private static void Main(string[] args)
        {
            try
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
                bool includeValid = false;

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
                                case "-x":
                                    returnXml = true;
                                    break;
                                case "--recursive":
                                    recursive = true;
                                    break;
                                case "-r":
                                    recursive = true;
                                    break;
                                case "--all":
                                    includeValid = true;
                                    break;
                                case "-a":
                                    includeValid = true;
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

                object validationErrors = validate.OOXML(xmlPath, version, returnXml, recursive, includeValid);

                Console.Write(validationErrors);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
