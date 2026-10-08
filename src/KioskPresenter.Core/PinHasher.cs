using System.Security.Cryptography;

namespace KioskPresenter.Core;

/// <summary>
/// Admin PIN storage. Only a salted PBKDF2-SHA256 hash is ever written to disk.
/// Format: pbkdf2-sha256$iterations$saltBase64$hashBase64
/// </summary>
public static class PinHasher
{
    public const int MinLength = 4;
    public const int MaxLength = 8;
    private const int Iterations = 200_000;
    private const int SaltBytes = 16;
    private const int HashBytes = 32;
    private const string Scheme = "pbkdf2-sha256";

    public static bool IsValidPin(string? pin) =>
        pin is not null && pin.Length is >= MinLength and <= MaxLength && pin.All(char.IsAsciiDigit);

    public static string Hash(string pin)
    {
        if (!IsValidPin(pin)) throw new ArgumentException($"PIN must be {MinLength} to {MaxLength} digits.", nameof(pin));
        var salt = RandomNumberGenerator.GetBytes(SaltBytes);
        var hash = Rfc2898DeriveBytes.Pbkdf2(pin, salt, Iterations, HashAlgorithmName.SHA256, HashBytes);
        return $"{Scheme}${Iterations}${Convert.ToBase64String(salt)}${Convert.ToBase64String(hash)}";
    }

    public static bool Verify(string? pin, string? stored)
    {
        if (pin is null || string.IsNullOrEmpty(stored)) return false;
        var parts = stored.Split('$');
        if (parts.Length != 4 || parts[0] != Scheme) return false;
        if (!int.TryParse(parts[1], out var iterations) || iterations < 1) return false;
        byte[] salt, expected;
        try
        {
            salt = Convert.FromBase64String(parts[2]);
            expected = Convert.FromBase64String(parts[3]);
        }
        catch (FormatException)
        {
            return false;
        }
        var actual = Rfc2898DeriveBytes.Pbkdf2(pin, salt, iterations, HashAlgorithmName.SHA256, expected.Length);
        return CryptographicOperations.FixedTimeEquals(actual, expected);
    }
}
