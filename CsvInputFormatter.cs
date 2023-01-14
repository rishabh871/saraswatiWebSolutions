using Core.Application.Extensions;
using Core.Application.MetaData;
using Microsoft.AspNetCore.Mvc.Formatters;
using Microsoft.Net.Http.Headers;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace API.Formatters
{
    public class CsvInputFormatter : InputFormatter
    {
        private readonly CsvFormatterOptions _options;

        public CsvInputFormatter(CsvFormatterOptions csvFormatterOptions)
        {
            SupportedMediaTypes.Add(MediaTypeHeaderValue.Parse("text/csv"));
            _options = csvFormatterOptions ?? throw new ArgumentNullException(nameof(csvFormatterOptions));
        }

        public override Task<InputFormatterResult> ReadRequestBodyAsync(InputFormatterContext context)
        {
            var type = context.ModelType;
            var request = context.HttpContext.Request;
            MediaTypeHeaderValue.TryParse(request.ContentType, out MediaTypeHeaderValue requestContentType);

            //var path = Path.GetTempFileName();
            //using (var stream = new FileStream(path, FileMode.Create))
            //{
            //    request.Body.CopyToAsync(stream);
            //}

            var result = ReadStream(type, request.Body);
            return InputFormatterResult.SuccessAsync(result);
        }

        public override bool CanRead(InputFormatterContext context)
        {
            var type = context.ModelType;
            if (type == null)
                throw new ArgumentNullException("type");

            return IsTypeOfIEnumerable(type);
        }

        private bool IsTypeOfIEnumerable(Type type)
        {

            foreach (Type interfaceType in type.GetInterfaces())
            {

                if (interfaceType == typeof(IList))
                    return true;
            }

            return false;
        }

        protected object ReadStream(Type type, Stream stream)
        {
            var reader = new StreamReader(stream, Encoding.GetEncoding(_options.Encoding));
            var csvContent = reader.ReadToEnd();
            if (string.IsNullOrEmpty(csvContent)) return null;
            bool skipFirstLine = _options.UseSingleLineHeaderInCsv;
            return csvContent.ToListOfType(type);
        }
    }
}
