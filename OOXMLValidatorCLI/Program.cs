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
            string xmlPath = string.Empty;
            string version = null;

            if (args != null && 0 < args.Length)
            {
                xmlPath = args[0];

                if (1 < args.Length)
                {
                    version = args[1];
                }
            }
            else
            {
                throw new ArgumentNullException();
            }

            bool returnXml = args[2] == "--xml" ? true : false;

            object validationErrors = validate.OOXML(xmlPath, version, returnXml);

            Console.Write(validationErrors);
        }
    }
}
