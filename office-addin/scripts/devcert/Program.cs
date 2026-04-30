using System.Net;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

if (args.Length != 3)
{
    Console.Error.WriteLine("Usage: DevCert <pfxPath> <cerPath> <passPath>");
    return 1;
}

var pfxPath = args[0];
var cerPath = args[1];
var passPath = args[2];

Directory.CreateDirectory(Path.GetDirectoryName(pfxPath)!);

var password = File.Exists(passPath)
    ? File.ReadAllText(passPath).Trim()
    : Guid.NewGuid().ToString("N");

File.WriteAllText(passPath, password);

using var rsa = RSA.Create(2048);
var request = new CertificateRequest(
    "CN=localhost",
    rsa,
    HashAlgorithmName.SHA256,
    RSASignaturePadding.Pkcs1);

var san = new SubjectAlternativeNameBuilder();
san.AddDnsName("localhost");
san.AddIpAddress(IPAddress.Loopback);
request.CertificateExtensions.Add(san.Build());
request.CertificateExtensions.Add(new X509EnhancedKeyUsageExtension(
    [new Oid("1.3.6.1.5.5.7.3.1", "Server Authentication")],
    critical: false));
request.CertificateExtensions.Add(new X509KeyUsageExtension(
    X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.KeyEncipherment,
    critical: true));
request.CertificateExtensions.Add(new X509BasicConstraintsExtension(
    certificateAuthority: false,
    hasPathLengthConstraint: false,
    pathLengthConstraint: 0,
    critical: true));

var notBefore = DateTimeOffset.UtcNow.AddMinutes(-5);
var notAfter = DateTimeOffset.UtcNow.AddYears(3);
using var certificate = request.CreateSelfSigned(notBefore, notAfter);
var exportable = new X509Certificate2(
    certificate.Export(X509ContentType.Pfx, password),
    password,
    X509KeyStorageFlags.Exportable | X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.UserKeySet);

File.WriteAllBytes(pfxPath, exportable.Export(X509ContentType.Pfx, password));
File.WriteAllBytes(cerPath, exportable.Export(X509ContentType.Cert));

using var rootStore = new X509Store(StoreName.Root, StoreLocation.CurrentUser);
rootStore.Open(OpenFlags.ReadWrite);
var existing = rootStore.Certificates.Find(X509FindType.FindByThumbprint, exportable.Thumbprint, validOnly: false);
if (existing.Count == 0)
{
    rootStore.Add(exportable);
}

Console.WriteLine(exportable.Thumbprint);
return 0;
