using DocumentFormat.OpenXml.Validation;
using Newtonsoft.Json;
using OOXMLValidator;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;

namespace OOXMLValidatorCLI
{
    class Program
    {
        static void Main(string[] args)
        {
            int? version = null;
            if (1 < args.Length && args[1] != null && int.TryParse(args[1], out int v))
            {
                version = v;
            }
            IEnumerable<ValidationErrorInfo> errors = Validate.OOXML(args[0], version);
            List<dynamic> res = new List<dynamic>();
            foreach (ValidationErrorInfo validationErrorInfo in errors)
            {
                dynamic dyno = new ExpandoObject();
                dyno.Description = validationErrorInfo.Description;
                dyno.Path = validationErrorInfo.Path;
                dyno.Id = validationErrorInfo.Id;
                dyno.ErrorType = validationErrorInfo.ErrorType;
                res.Add(dyno);
            }
            var json = JsonConvert.SerializeObject(res, Formatting.None,
                        new JsonSerializerSettings()
                        {
                            ReferenceLoopHandling = ReferenceLoopHandling.Ignore
                        });
            Console.WriteLine(json);
        }
    }
}
