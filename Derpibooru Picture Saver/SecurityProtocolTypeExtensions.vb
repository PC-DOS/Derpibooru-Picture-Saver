Public Module SecurityProtocolTypeExtensionsModule
    Public Enum SecurituProctocolTypeExtensions
        Tls11 = 768
        Tls12 = 3072
        Tls13 = 12288 'Currently not supported, see https://docs.microsoft.com/zh-cn/dotnet/api/system.net.securityprotocoltype?view=net-5.0
    End Enum

End Module
