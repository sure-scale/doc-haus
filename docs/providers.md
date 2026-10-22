# Model providers

Doc Haus runs model requests and connection checks through its native OpenCode
engine. OpenCode is the execution foundation; it does not need to be selected as
a hosted model provider.

The hosted provider named **OpenCode** (`opencode`, including its public free
models) is unavailable in Doc Haus. It is omitted from provider, model and saved
sign-in choices and cannot be re-enabled through installation settings.
**OpenCode Go** (`opencode-go`) remains a separate supported provider, alongside
the other supported hosted and local providers.

If an existing installation selected an `opencode/…` model, open **Settings →
Providers** to connect another provider, then choose and save its primary and
fast models under **Models**. Doc Haus asks for a deliberate replacement; it
does not silently select another provider or convert a saved sign-in to Go.
Existing credentials and conversations remain stored.

Connection checks send a short fixed prompt through native model resolution and
execution. They do not send matter text or resume drafting. A successful check
confirms the provider accepted that request; normal legal workflows retain their
own permissions and review steps.
