using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Text.Json;
using System.Threading.Tasks;

namespace MyExcelMAUIApp
{
    public class EntryConverter : JsonConverter<Entry>
    {
        public override Entry Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            if (reader.TokenType == JsonTokenType.String)
            {
                string text = reader.GetString();
                return new Entry { Text = text };
            }

            return null; 
        }

        public override void Write(Utf8JsonWriter writer, Entry value, JsonSerializerOptions options)
        {
            writer.WriteStringValue(value.Text);
        }
    }
}
