using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using DocumentFormat.OpenXml.Validation;
using Microsoft.Extensions.DependencyInjection;
using Newtonsoft.Json;
using OOXMLValidatorCLI;
using OOXMLValidatorCLI.Classes;
using OOXMLValidatorCLI.Interfaces;

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
            collection.AddScoped<IDocument, Document>();
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

            string json = validate.OOXML(xmlPath, version);

            Console.Write(json);
        }
    }
}
