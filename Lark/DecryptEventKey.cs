using System.Security.Cryptography;

namespace Elfland.Desert.Lark;

public static class DecryptEventKey
{
    public static string CalculateSignature(
        string timestamp,
        string nonce,
        string encryptKey,
        string body
    )
    {
        StringBuilder content = new StringBuilder();
        content.Append(timestamp);
        content.Append(nonce);
        content.Append(encryptKey);
        content.Append(body);
        var sha256 = SHA256.Create();
        var bytes_out = sha256.ComputeHash(Encoding.Default.GetBytes(content.ToString()));
        var result = BitConverter.ToString(bytes_out);
        return result.Replace("-", "");
    }
}
