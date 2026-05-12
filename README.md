# ULS-EmailGenerator

Tools for generating Flex Project emails on `portal.ul.com/Project/Details/*`
from local `.xlsx` templates (typically synced via OneDrive).

Two equivalent distributions are provided:

| Distribution | When to use it | Entry point |
|---|---|---|
| **Tampermonkey userscript** | You already use Tampermonkey / Violentmonkey and want a single-file install. | [`FlexEmailGeneratorExcelBased.user.js`](FlexEmailGeneratorExcelBased.user.js) |
| **Chrome / Edge extension** (MV3) | You want a regular installable extension with a toolbar popup, no userscript manager required. | [`extension/`](extension/) — see [`extension/README.md`](extension/README.md) for install steps |

Both share the same template engine and the same Excel sheet schema
(documented in the long comment header at the top of the userscript).
The extension's `content/main.js` is a near-verbatim copy of the
userscript; the only differences are a small `gm-compat.js` shim
that maps `GM_*` / `unsafeWindow` onto standard Web Extension APIs,
and a toolbar popup that replaces `GM_registerMenuCommand`.

See [`extension/README.md`](extension/README.md) for detailed install
and usage instructions.
