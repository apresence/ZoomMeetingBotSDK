// Snarfed from https://github.com/dotnet/runtime/issues/31024

namespace ZoomMeetingBotSDK
{
    using System;
    using System.Text.Json;
    using System.Text.Json.Serialization;

    internal class JSONSpecialDoubleHandler
    {
        public class HandleSpecialDoublesAsStrings : JsonConverter<double>
        {
            public override double Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
            {
                if (reader.TokenType == JsonTokenType.String)
                {
                    return double.Parse(reader.GetString());
                }
                return reader.GetDouble();
            }

            public override void Write(Utf8JsonWriter writer, double value, JsonSerializerOptions options)
            {
                // IsFinite is missing in 4.8, so we have to check !IsInfinity instead
                if (!double.IsInfinity(value))
                {
                    writer.WriteNumberValue(value);
                }
                else
                {
                    writer.WriteStringValue(value.ToString());
                }
            }

            public static string Serialize(object value)
            {
                var options = new JsonSerializerOptions();
                options.Converters.Add(new HandleSpecialDoublesAsStrings());
                // Don't be hyper-aggresive with escaping, i.e. don't escape "'"
                options.Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping;
                return JsonSerializer.Serialize(value, options);
            }
        }

        private class HandleSpecialDoublesAsStrings_NewtonsoftCompat : JsonConverter<double>
        {
            public override double Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
            {
                if (reader.TokenType == JsonTokenType.String)
                {
                    string specialDouble = reader.GetString();
                    if (specialDouble == "Infinity")
                    {
                        return double.PositiveInfinity;
                    }
                    else if (specialDouble == "-Infinity")
                    {
                        return double.NegativeInfinity;
                    }
                    else
                    {
                        return double.NaN;
                    }
                }
                return reader.GetDouble();
            }

            private static readonly JsonEncodedText S_nan = JsonEncodedText.Encode("NaN");
            private static readonly JsonEncodedText S_infinity = JsonEncodedText.Encode("Infinity");
            private static readonly JsonEncodedText S_negativeInfinity = JsonEncodedText.Encode("-Infinity");

            public override void Write(Utf8JsonWriter writer, double value, JsonSerializerOptions options)
            {
                // IsFinite is missing in 4.8
                if (!double.IsInfinity(value))
                {
                    writer.WriteNumberValue(value);
                }
                else
                {
                    if (double.IsPositiveInfinity(value))
                    {
                        writer.WriteStringValue(S_infinity);
                    }
                    else if (double.IsNegativeInfinity(value))
                    {
                        writer.WriteStringValue(S_negativeInfinity);
                    }
                    else
                    {
                        writer.WriteStringValue(S_nan);
                    }
                }
            }
        }
    }
}
