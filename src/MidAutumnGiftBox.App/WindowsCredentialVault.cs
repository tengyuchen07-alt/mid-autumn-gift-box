using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using MidAutumnGiftBox.Core;

namespace MidAutumnGiftBox.App;

internal sealed class WindowsCredentialVault : ICredentialVault
{
    private const uint CredentialTypeGeneric = 1;
    private const uint CredentialPersistLocalMachine = 2;
    private const int MaximumCredentialBlobBytes = 2_560;
    private const string TargetPrefix = "MidAutumnGiftBox/Wms/";

    public Task SaveAsync(
        string sourceCode,
        ApiCredentials credentials,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var secretBytes = Encoding.Unicode.GetBytes(credentials.ApiKey);
        if (secretBytes.Length > MaximumCredentialBlobBytes)
        {
            throw new ArgumentException("API Key 長度超過 Windows Credential Manager 限制。", nameof(credentials));
        }

        var secretPointer = Marshal.AllocCoTaskMem(secretBytes.Length);
        try
        {
            Marshal.Copy(secretBytes, 0, secretPointer, secretBytes.Length);
            var credential = new NativeCredential
            {
                Type = CredentialTypeGeneric,
                TargetName = TargetPrefix + sourceCode,
                CredentialBlobSize = (uint)secretBytes.Length,
                CredentialBlob = secretPointer,
                Persist = CredentialPersistLocalMachine,
                UserName = credentials.ApiId
            };

            if (!CredWrite(ref credential, 0))
            {
                throw new Win32Exception(Marshal.GetLastWin32Error(), "無法安全保存 API 憑證。");
            }
        }
        finally
        {
            if (secretBytes.Length > 0)
            {
                Marshal.Copy(new byte[secretBytes.Length], 0, secretPointer, secretBytes.Length);
            }

            Marshal.FreeCoTaskMem(secretPointer);
            Array.Clear(secretBytes);
        }

        return Task.CompletedTask;
    }

    public Task<ApiCredentials?> LoadAsync(
        string sourceCode,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (!CredRead(TargetPrefix + sourceCode, CredentialTypeGeneric, 0, out var credentialPointer))
        {
            const int ErrorNotFound = 1168;
            var error = Marshal.GetLastWin32Error();
            if (error == ErrorNotFound)
            {
                return Task.FromResult<ApiCredentials?>(null);
            }

            throw new Win32Exception(error, "無法讀取已保存的 API 憑證。");
        }

        try
        {
            var credential = Marshal.PtrToStructure<NativeCredential>(credentialPointer);
            var secret = credential.CredentialBlobSize == 0
                ? string.Empty
                : Marshal.PtrToStringUni(credential.CredentialBlob, (int)credential.CredentialBlobSize / 2) ?? string.Empty;
            return Task.FromResult<ApiCredentials?>(new ApiCredentials(credential.UserName ?? string.Empty, secret));
        }
        finally
        {
            CredFree(credentialPointer);
        }
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct NativeCredential
    {
        public uint Flags;
        public uint Type;
        public string? TargetName;
        public string? Comment;
        public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
        public uint CredentialBlobSize;
        public IntPtr CredentialBlob;
        public uint Persist;
        public uint AttributeCount;
        public IntPtr Attributes;
        public string? TargetAlias;
        public string? UserName;
    }

    [DllImport("advapi32.dll", EntryPoint = "CredWriteW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool CredWrite(ref NativeCredential credential, uint flags);

    [DllImport("advapi32.dll", EntryPoint = "CredReadW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool CredRead(string target, uint type, uint reservedFlag, out IntPtr credentialPointer);

    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern void CredFree(IntPtr credentialPointer);
}
