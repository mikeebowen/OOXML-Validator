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
            var collection = new ServiceCollection();
            collection.AddScoped<IValidate, Validate>();
            collection.AddScoped<IFunctionUtils, FunctionUtils>();
            collection.AddScoped<IDocumentUtils, DocumentUtils>();
            var serviceProvider = collection.BuildServiceProvider();

            var validate = serviceProvider.GetService<IValidate>();
            string xmlPath;
            bool? returnXml = null;
            string version = null;

            if (args != null && 0 < args.Length)
            {
                xmlPath = args[0];

                if (args.Length == 2)
                {
                    version = args[1];
                    returnXml = args[1] == "--xml" ? true : false;
                }

                if (args.Length == 3)
                {
                    if (!args[1].StartsWith("--"))
                    {
                        version = args[1];
                    }
                    else
                    {
                        version = args[2];
                    }

                    returnXml = args[1] == "--xml" || args[2] == "--xml" ? true : false;
                }

                if (args.Length > 3)
                {
                    throw new ArgumentOutOfRangeException(nameof(args));
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

            object validationErrors = validate.OOXML(xmlPath, version, returnXml ?? false);

            Console.Write(validationErrors);
        }
    }
}
