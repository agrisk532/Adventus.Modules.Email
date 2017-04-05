using System;
using System.Text;

namespace Adventus.Modules.Email
{
	public class Encoder
	{

	//	from http://stackoverflow.com/questions/11793734/code-for-encode-decode-quotedprintable

		public static string EncodeQuotedPrintable(string value)
		{
		    if (string.IsNullOrEmpty(value))
		        return value;
		
		    StringBuilder builder = new StringBuilder();
		
		    byte[] bytes = Encoding.UTF8.GetBytes(value);
		    foreach (byte v in bytes)
		    {
		        // The following are not required to be encoded:
		        // - Tab (ASCII 9)
		        // - Space (ASCII 32)
		        // - Characters 33 to 126, except for the equal sign (61).
		
		        if ((v == 9) || ((v >= 32) && (v <= 60)) || ((v >= 62) && (v <= 126)))
		        {
		            builder.Append(Convert.ToChar(v));
		        }
		        else
		        {
		            builder.Append('=');
		            builder.Append(v.ToString("X2"));
		        }
		    }
		
		    char lastChar = builder[builder.Length - 1];
		    if (char.IsWhiteSpace(lastChar))
		    {
		        builder.Remove(builder.Length - 1, 1);
		        builder.Append('=');
		        builder.Append(((int)lastChar).ToString("X2"));
		    }
		
		    return builder.ToString();
		}
	}
}
