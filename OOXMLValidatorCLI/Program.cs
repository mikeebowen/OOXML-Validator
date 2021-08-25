using DocumentFormat.OpenXml.Validation;
using Microsoft.Extensions.DependencyInjection;
using Newtonsoft.Json;
using OOXMLValidator;
using OOXMLValidator.Classes;
using OOXMLValidator.Interfaces;
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
            collection.AddScoped<IValidate, OOXMLValidator.Classes.Validate>();
            collection.AddScoped<IFunctionUtils, FunctionUtils>();
            collection.AddScoped<IDocument, Document>();
            var serviceProvider = collection.BuildServiceProvider();

            int? version = null;
            if (1 < args.Length && args[1] != null && int.TryParse(args[1], out int v))
            {
                version = v;
            }

            var validate = serviceProvider.GetService<IValidate>();
            string json = validate.OOXML(args[0], version);

            Console.Write(json);
        }
    }
}
