namespace MidAutumnGiftBox.Core;

public interface ICredentialVault
{
    Task SaveAsync(
        string sourceCode,
        ApiCredentials credentials,
        CancellationToken cancellationToken = default);

    Task<ApiCredentials?> LoadAsync(
        string sourceCode,
        CancellationToken cancellationToken = default);
}

public sealed record CredentialDisplayStatus(string ApiId, bool HasSavedKey);

public sealed class CredentialManager(ICredentialVault vault)
{
    public async Task SaveAsync(
        string sourceCode,
        ApiCredentials credentials,
        CancellationToken cancellationToken = default)
    {
        ValidateSourceCode(sourceCode);
        if (!credentials.IsComplete)
        {
            throw new ArgumentException("請輸入 API ID 與 API Key。", nameof(credentials));
        }

        await vault.SaveAsync(sourceCode, credentials, cancellationToken);
    }

    public async Task<ApiCredentials?> GetForUseAsync(
        string sourceCode,
        CancellationToken cancellationToken = default)
    {
        ValidateSourceCode(sourceCode);
        return await vault.LoadAsync(sourceCode, cancellationToken);
    }

    public async Task<CredentialDisplayStatus> GetDisplayStatusAsync(
        string sourceCode,
        CancellationToken cancellationToken = default)
    {
        var credentials = await GetForUseAsync(sourceCode, cancellationToken);
        return credentials is null
            ? new CredentialDisplayStatus(string.Empty, false)
            : new CredentialDisplayStatus(credentials.ApiId, true);
    }

    public async Task<ApiCredentials> ResolveForUseAsync(
        string sourceCode,
        string displayedApiId,
        string enteredApiKey,
        CancellationToken cancellationToken = default)
    {
        ValidateSourceCode(sourceCode);
        var normalizedApiId = displayedApiId.Trim();
        if (!string.IsNullOrWhiteSpace(enteredApiKey))
        {
            var entered = new ApiCredentials(normalizedApiId, enteredApiKey);
            if (!entered.IsComplete)
            {
                throw new InvalidOperationException("請輸入 API ID 與 API Key。");
            }

            return entered;
        }

        var stored = await vault.LoadAsync(sourceCode, cancellationToken)
            ?? throw new InvalidOperationException("請先輸入 API ID 與 API Key，或保存這個來源的憑證。");
        if (!string.Equals(stored.ApiId, normalizedApiId, StringComparison.Ordinal))
        {
            throw new InvalidOperationException("畫面上的 API ID 已變更，請同時輸入新的 API Key，再重新測試連線。");
        }

        return stored;
    }

    private static void ValidateSourceCode(string sourceCode)
    {
        if (string.IsNullOrWhiteSpace(sourceCode) ||
            sourceCode.Any(character => !char.IsAsciiLetterOrDigit(character) && character is not '-' and not '_'))
        {
            throw new ArgumentException("來源代碼格式不正確。", nameof(sourceCode));
        }
    }
}
