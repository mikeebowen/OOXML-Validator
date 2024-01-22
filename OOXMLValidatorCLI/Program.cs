// Licensed under the MIT license. See LICENSE file in the project root for full license information.

OOXMLValidatorCLI.Program.Start(args);

namespace OOXMLValidatorCLI
{
    using System;
    using DocumentFormat.OpenXml;
    using Microsoft.Extensions.DependencyInjection;
    using OOXMLValidatorCLI.Classes;
    using OOXMLValidatorCLI.Interfaces;

    /// <summary>
    /// Represents the program's entry point.
    /// </summary>
    public class Program
    {
        /// <summary>
        /// Entry point of the application.
        /// </summary>
        /// <param name="args">Command line arguments.</param>
        public static void Start(string[] args)
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

                if (args is not null && args.Length > 0)
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
