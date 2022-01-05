using DocumentFormat.OpenXml.Validation;
using Microsoft.Extensions.DependencyInjection;
using Newtonsoft.Json;
using OOXMLValidatorCLI;
using OOXMLValidatorCLI.Classes;
using OOXMLValidatorCLI.Interfaces;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;

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
            string json = validate.OOXML(args[0], args[1]);

            Console.Write(json);
        }
    }
}
