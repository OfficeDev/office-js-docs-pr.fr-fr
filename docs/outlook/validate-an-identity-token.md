---
title: Valider un jeton d’identité de complément Outlook
description: Votre complément Outlook peut vous envoyer un jeton d’identité d’utilisateur Exchange, mais avant de faire confiance à la requête, vous devez valider le jeton pour vous assurer qu’il provient du serveur Exchange attendu.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 6ad5f99093530528ec83cfc7a6e3a2571e0df491
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094105"
---
# <a name="validate-an-exchange-identity-token"></a><span data-ttu-id="a675b-103">Valider un jeton d’identité Exchange</span><span class="sxs-lookup"><span data-stu-id="a675b-103">Validate an Exchange identity token</span></span>

<span data-ttu-id="a675b-104">Votre complément Outlook peut vous envoyer un jeton d’identité d’utilisateur Exchange, mais avant de faire confiance à la requête, vous devez valider le jeton pour vous assurer qu’il provient du serveur Exchange attendu.</span><span class="sxs-lookup"><span data-stu-id="a675b-104">Your Outlook add-in can send you an Exchange user identity token, but before you trust the request you must validate the token to ensure that it came from the Exchange server that you expect.</span></span> <span data-ttu-id="a675b-105">Les jetons d’identité d’utilisateur Exchange sont des jetons Web JSON (JWT).</span><span class="sxs-lookup"><span data-stu-id="a675b-105">Exchange user identity tokens are JSON Web Tokens (JWT).</span></span> <span data-ttu-id="a675b-106">Les étapes nécessaires pour valider un jeton JWT sont décrites dans le document [RFC 7519 JSON Web Token (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt).</span><span class="sxs-lookup"><span data-stu-id="a675b-106">The steps required to validate a JWT are described in [RFC 7519 JSON Web Token (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt).</span></span>

<span data-ttu-id="a675b-107">Nous vous suggérons d’utiliser un processus en quatre étapes pour valider le jeton d’identité et obtenir l’identificateur unique de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="a675b-107">We suggest that you use a four-step process to validate the identity token and obtain the user's unique identifier.</span></span> <span data-ttu-id="a675b-108">Dans un premier temps, extrayez le jeton JWT (JSON Web Token) à partir d’une chaîne d’URL encodée au format base64.</span><span class="sxs-lookup"><span data-stu-id="a675b-108">First, extract the JSON Web Token (JWT) from a base64 URL-encoded string.</span></span> <span data-ttu-id="a675b-109">Dans un deuxième temps, assurez-vous que le jeton est bien formé, c’est-à-dire qu’il est adapté à votre complément Outlook, qu’il n’a pas expiré et que vous pouvez extraire une URL valide pour le document de métadonnées d’authentification.</span><span class="sxs-lookup"><span data-stu-id="a675b-109">Second, make sure that the token is well-formed, that it is for your Outlook add-in, that it has not expired, and that you can extract a valid URL for the authentication metadata document.</span></span> <span data-ttu-id="a675b-110">Dans un troisième temps, récupérez le document de métadonnées d’authentification sur le serveur Exchange et validez la signature jointe au jeton d’identité.</span><span class="sxs-lookup"><span data-stu-id="a675b-110">Next, retrieve the authentication metadata document from the Exchange server and validate the signature attached to the identity token.</span></span> <span data-ttu-id="a675b-111">Enfin, calculez un identificateur unique pour l’utilisateur en concaténant l’ID Exchange de l’utilisateur avec l’URL du document de métadonnées d’authentification.</span><span class="sxs-lookup"><span data-stu-id="a675b-111">Finally, compute a unique identifier for the user by concatenating the user's Exchange ID with the URL of the authentication metadata document.</span></span>

## <a name="extract-the-json-web-token"></a><span data-ttu-id="a675b-112">Extraction du jeton Web JSON</span><span class="sxs-lookup"><span data-stu-id="a675b-112">Extract the JSON Web Token</span></span>

<span data-ttu-id="a675b-113">Le jeton renvoyé par [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) est une représentation de chaîne encodée du jeton.</span><span class="sxs-lookup"><span data-stu-id="a675b-113">The token returned from [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) is an encoded string representation of the token.</span></span> <span data-ttu-id="a675b-114">Dans ce formulaire, conformément au document RFC 7519, tous les jetons JWT se composent de trois parties, séparées par un point.</span><span class="sxs-lookup"><span data-stu-id="a675b-114">In this form, per RFC 7519, all JWTs have three parts, separated by a period.</span></span> <span data-ttu-id="a675b-115">Le format est comme suit.</span><span class="sxs-lookup"><span data-stu-id="a675b-115">The format is as follows.</span></span>

```json
{header}.{payload}.{signature}
```

<span data-ttu-id="a675b-116">L’en-tête et la charge utile doivent être décodés au format base64 pour obtenir une représentation JSON de chaque partie.</span><span class="sxs-lookup"><span data-stu-id="a675b-116">The header and payload should be base64-decoded to obtain a JSON representation of each part.</span></span> <span data-ttu-id="a675b-117">La signature doit être décodée au format base64 pour obtenir un tableau d’octets contenant la signature binaire.</span><span class="sxs-lookup"><span data-stu-id="a675b-117">The signature should be base64-decoded to obtain a byte array containing the binary signature.</span></span>

<span data-ttu-id="a675b-118">Pour plus d’informations sur le contenu du jeton, consultez la section [Présentation du jeton d’identité Exchange](inside-the-identity-token.md).</span><span class="sxs-lookup"><span data-stu-id="a675b-118">For more information about the contents of the token, see [Inside the Exchange identity token](inside-the-identity-token.md).</span></span>

<span data-ttu-id="a675b-119">Une fois les trois composants décodés, vous pouvez poursuivre avec la validation du contenu du jeton.</span><span class="sxs-lookup"><span data-stu-id="a675b-119">After you have the three decoded components, you can proceed with validating the content of the token.</span></span>

## <a name="validate-token-contents"></a><span data-ttu-id="a675b-120">Validation du contenu du jeton</span><span class="sxs-lookup"><span data-stu-id="a675b-120">Validate token contents</span></span>

<span data-ttu-id="a675b-121">Pour valider le contenu du jeton, vous devez vérifier ce qui suit.</span><span class="sxs-lookup"><span data-stu-id="a675b-121">To validate the token contents, you should check the following.</span></span>

- <span data-ttu-id="a675b-122">Vérifiez l’en-tête et assurez-vous que :</span><span class="sxs-lookup"><span data-stu-id="a675b-122">Check the header and verify that the:</span></span>
    - <span data-ttu-id="a675b-123">`typ`la revendication est définie sur `JWT` .</span><span class="sxs-lookup"><span data-stu-id="a675b-123">`typ` claim is set to `JWT`.</span></span>
    - <span data-ttu-id="a675b-124">`alg`la revendication est définie sur `RS256` .</span><span class="sxs-lookup"><span data-stu-id="a675b-124">`alg` claim is set to `RS256`.</span></span>
    - <span data-ttu-id="a675b-125">`x5t`la revendication est présente.</span><span class="sxs-lookup"><span data-stu-id="a675b-125">`x5t` claim is present.</span></span>

- <span data-ttu-id="a675b-126">Vérifiez la charge utile et assurez-vous que :</span><span class="sxs-lookup"><span data-stu-id="a675b-126">Check the payload and verify that the:</span></span>
    - <span data-ttu-id="a675b-127">`amurl`la revendication dans le `appctx` est définie sur l’emplacement d’un fichier manifeste de clés de signature de jeton autorisé.</span><span class="sxs-lookup"><span data-stu-id="a675b-127">`amurl` claim inside the `appctx` is set to the location of an authorized token signing key manifest file.</span></span> <span data-ttu-id="a675b-128">Par exemple, la valeur attendue `amurl` pour Microsoft 365 est https://outlook.office365.com:443/autodiscover/metadata/json/1 .</span><span class="sxs-lookup"><span data-stu-id="a675b-128">For example, the expected `amurl` value for Microsoft 365 is https://outlook.office365.com:443/autodiscover/metadata/json/1.</span></span> <span data-ttu-id="a675b-129">Pour plus d’informations, reportez-vous [à la section](#verify-the-domain) suivante.</span><span class="sxs-lookup"><span data-stu-id="a675b-129">See the next section [Verify the domain](#verify-the-domain) for additional information.</span></span>
    - <span data-ttu-id="a675b-130">L’heure actuelle est comprise entre les heures spécifiées dans les `nbf` `exp` revendications et.</span><span class="sxs-lookup"><span data-stu-id="a675b-130">Current time is between the times specified in the `nbf` and `exp` claims.</span></span> <span data-ttu-id="a675b-131">La revendication `nbf` spécifie le début de la période où le jeton est considéré comme valide et la revendication `exp` spécifie le délai d’expiration pour le jeton.</span><span class="sxs-lookup"><span data-stu-id="a675b-131">The `nbf` claim specifies the earliest time that the token is considered valid, and the `exp` claim specifies the expiration time for the token.</span></span> <span data-ttu-id="a675b-132">Ceci est recommandé pour permettre certains écarts dans les paramètres de l’horloge entre les serveurs.</span><span class="sxs-lookup"><span data-stu-id="a675b-132">It is recommended to allow for some variation in clock settings between servers.</span></span>
    - <span data-ttu-id="a675b-133">`aud`claim est l’URL attendue pour votre complément.</span><span class="sxs-lookup"><span data-stu-id="a675b-133">`aud` claim is the expected URL for your add-in.</span></span>
    - <span data-ttu-id="a675b-134">`version`la revendication à l’intérieur de la `appctx` revendication est définie sur `ExIdTok.V1` .</span><span class="sxs-lookup"><span data-stu-id="a675b-134">`version` claim inside the `appctx` claim is set to `ExIdTok.V1`.</span></span>

### <a name="verify-the-domain"></a><span data-ttu-id="a675b-135">Vérifier le domaine</span><span class="sxs-lookup"><span data-stu-id="a675b-135">Verify the domain</span></span>

<span data-ttu-id="a675b-136">Lors de l’implémentation de la logique de vérification décrite précédemment dans cette section, vous devez également exiger que le domaine de la `amurl` revendication corresponde au domaine de découverte automatique de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="a675b-136">When implementing the verification logic described previously in this section, you should also require that the domain of the `amurl` claim matches the Autodiscover domain for the user.</span></span> <span data-ttu-id="a675b-137">Pour ce faire, vous devez utiliser ou implémenter la découverte automatique.</span><span class="sxs-lookup"><span data-stu-id="a675b-137">To do so, you'll need to use or implement Autodiscover.</span></span> <span data-ttu-id="a675b-138">Pour en savoir plus, vous pouvez commencer à utiliser la [découverte automatique pour Exchange](/exchange/client-developer/exchange-web-services/autodiscover-for-exchange).</span><span class="sxs-lookup"><span data-stu-id="a675b-138">To learn more, you can start with [Autodiscover for Exchange](/exchange/client-developer/exchange-web-services/autodiscover-for-exchange).</span></span>

## <a name="validate-the-identity-token-signature"></a><span data-ttu-id="a675b-139">Validation de la signature du jeton d’identité</span><span class="sxs-lookup"><span data-stu-id="a675b-139">Validate the identity token signature</span></span>

<span data-ttu-id="a675b-140">Une fois que vous savez que le jeton JWT contient les revendications requises, poursuivez avec la validation de la signature du jeton.</span><span class="sxs-lookup"><span data-stu-id="a675b-140">After you know that the JWT contains the required claims, you can proceed with validating the token signature.</span></span>

### <a name="retrieve-the-public-signing-key"></a><span data-ttu-id="a675b-141">Récupération de la clé de signature publique</span><span class="sxs-lookup"><span data-stu-id="a675b-141">Retrieve the public signing key</span></span>

<span data-ttu-id="a675b-142">La première étape consiste à récupérer la clé publique qui correspond au certificat que le serveur Exchange a utilisé pour signer le jeton.</span><span class="sxs-lookup"><span data-stu-id="a675b-142">The first step is to retrieve the public key that corresponds to the certificate that the Exchange server used to sign the token.</span></span> <span data-ttu-id="a675b-143">La clé est disponible dans le document de métadonnées d’authentification.</span><span class="sxs-lookup"><span data-stu-id="a675b-143">The key is found in the authentication metadata document.</span></span> <span data-ttu-id="a675b-144">Ce document est un fichier JSON hébergé dans l’URL spécifiée dans la réclamation `amurl`.</span><span class="sxs-lookup"><span data-stu-id="a675b-144">This document is a JSON file hosted at the URL specified in the `amurl` claim.</span></span>

<span data-ttu-id="a675b-145">Le document de métadonnées d’authentification utilise le format suivant.</span><span class="sxs-lookup"><span data-stu-id="a675b-145">The authentication metadata document uses the following format.</span></span>

```json
{
    "id": "_70b34511-d105-4e2b-9675-39f53305bb01",
    "version": "1.0",
    "name": "Exchange",
    "realm": "*",
    "serviceName": "00000002-0000-0ff1-ce00-000000000000",
    "issuer": "00000002-0000-0ff1-ce00-000000000000@*",
    "allowedAudiences": [
        "00000002-0000-0ff1-ce00-000000000000@*"
    ],
    "keys": [
        {
            "usage": "signing",
            "keyinfo": {
                "x5t": "enh9BJrVPU5ijV1qjZjV-fL2bco"
            },
            "keyvalue": {
                "type": "x509Certificate",
                "value": "MIIHNTCC..."
            }
        }
    ],
    "endpoints": [
        {
            "location": "https://by2pr06mb2229.namprd06.prod.outlook.com:444/autodiscover/metadata/json/1",
            "protocol": "OAuth2",
            "usage": "metadata"
        }
    ]
}
```

<span data-ttu-id="a675b-146">Les clés de signature disponibles sont dans le tableau `keys`.</span><span class="sxs-lookup"><span data-stu-id="a675b-146">The available signing keys are in the `keys` array.</span></span> <span data-ttu-id="a675b-147">Sélectionnez la clé correcte en vérifiant que la valeur `x5t` dans la propriété `keyinfo` correspond à la valeur `x5t` dans l’en-tête du jeton.</span><span class="sxs-lookup"><span data-stu-id="a675b-147">Select the correct key by ensuring that the `x5t` value in the `keyinfo` property matches the `x5t` value in the header of the token.</span></span> <span data-ttu-id="a675b-148">La clé publique est à l’intérieur de la propriété `value` dans la propriété `keyvalue`. Elle est stockée sous la forme d’un tableau d’octets codé au format base64.</span><span class="sxs-lookup"><span data-stu-id="a675b-148">The public key is inside the `value` property in the `keyvalue` property, stored as a base64-encoded byte array.</span></span>

<span data-ttu-id="a675b-149">Une fois que vous avez trouvé la bonne clé publique, vérifiez la signature.</span><span class="sxs-lookup"><span data-stu-id="a675b-149">After you have the correct public key, verify the signature.</span></span> <span data-ttu-id="a675b-150">Les données signées correspondent aux deux premières parties du jeton codé, séparées par un point :</span><span class="sxs-lookup"><span data-stu-id="a675b-150">The signed data is the first two parts of the encoded token, separated by a period:</span></span>

```json
{header}.{payload}
```

## <a name="compute-the-unique-id-for-an-exchange-account"></a><span data-ttu-id="a675b-151">Calculer l’ID unique d’un compte Exchange</span><span class="sxs-lookup"><span data-stu-id="a675b-151">Compute the unique ID for an Exchange account</span></span>

<span data-ttu-id="a675b-152">Vous pouvez créer un identificateur unique pour un compte Exchange en concaténant l’URL du document de métadonnées d’authentification avec l’identificateur Exchange pour le compte.</span><span class="sxs-lookup"><span data-stu-id="a675b-152">You can create a unique identifier for an Exchange account by concatenating the authentication metadata document URL with the Exchange identifier for the account.</span></span> <span data-ttu-id="a675b-153">Lorsque vous avez cet identificateur unique, vous pouvez l’utiliser pour créer un système de connexion unique (SSO) pour le service Web de votre complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="a675b-153">When you have this unique identifier, you can use it to create a single sign-on (SSO) system for your Outlook add-in web service.</span></span> <span data-ttu-id="a675b-154">Pour plus d’informations sur l’utilisation de l’identificateur unique pour l’authentification unique, consultez la section [Authentifier un utilisateur avec un jeton d’identité pour Exchange](authenticate-a-user-with-an-identity-token.md).</span><span class="sxs-lookup"><span data-stu-id="a675b-154">For details about using the unique identifier for SSO, see [Authenticate a user with an identity token for Exchange](authenticate-a-user-with-an-identity-token.md).</span></span>

## <a name="use-a-library-to-validate-the-token"></a><span data-ttu-id="a675b-155">Utiliser une bibliothèque pour valider le jeton</span><span class="sxs-lookup"><span data-stu-id="a675b-155">Use a library to validate the token</span></span>

<span data-ttu-id="a675b-156">Il existe un certain nombre de bibliothèques qui permettent une analyse et une validation générales du jeton JWT.</span><span class="sxs-lookup"><span data-stu-id="a675b-156">There are a number of libraries that can do general JWT parsing and validation.</span></span> <span data-ttu-id="a675b-157">Microsoft fournit la `System.IdentityModel.Tokens.Jwt` bibliothèque qui peut être utilisée pour valider les jetons d’identité d’utilisateur Exchange.</span><span class="sxs-lookup"><span data-stu-id="a675b-157">Microsoft provides the `System.IdentityModel.Tokens.Jwt` library that can be used to validate Exchange user identity tokens.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a675b-158">Nous ne recommandons plus l’API managée des services Web Exchange, car le Microsoft.Exchange.WebServices.Auth.dll, bien que toujours disponible, est désormais obsolète et s’appuie sur des bibliothèques non prises en charge, telles que Microsoft.IdentityModel.Extensions.dll.</span><span class="sxs-lookup"><span data-stu-id="a675b-158">We no longer recommend the Exchange Web Services Managed API because the Microsoft.Exchange.WebServices.Auth.dll, though still available, is now obsolete and relies on unsupported libraries like Microsoft.IdentityModel.Extensions.dll.</span></span>

### <a name="systemidentitymodeltokensjwt"></a><span data-ttu-id="a675b-159">System.IdentityModel.Tokens.Jwt</span><span class="sxs-lookup"><span data-stu-id="a675b-159">System.IdentityModel.Tokens.Jwt</span></span>

<span data-ttu-id="a675b-160">La bibliothèque [System.IdentityModels.Tokens.Jwt](https://www.nuget.org/packages/System.IdentityModel.Tokens.Jwt) peut analyser le jeton et également effectuer la validation, même si vous devez analyser la réclamation `appctx` vous-même et récupérer la clé de signature publique.</span><span class="sxs-lookup"><span data-stu-id="a675b-160">The [System.IdentityModels.Tokens.Jwt](https://www.nuget.org/packages/System.IdentityModel.Tokens.Jwt) library can parse the token and also perform the validation, though you will need to parse the `appctx` claim yourself and retrieve the public signing key.</span></span>

```cs
// Load the encoded token
string encodedToken = "...";
JwtSecurityToken jwt = new JwtSecurityToken(encodedToken);

// Parse the appctx claim to get the auth metadata url
string authMetadataUrl = string.Empty;
var appctx = jwt.Claims.FirstOrDefault(claim => claim.Type == "appctx");
if (appctx != null)
{
    var AppContext = JsonConvert.DeserializeObject<ExchangeAppContext>(appctx.Value);

    // Token version check
    if (string.Compare(AppContext.Version, "ExIdTok.V1", StringComparison.InvariantCulture) != 0) {
        // Fail validation
    }

    authMetadataUrl = AppContext.MetadataUrl;
}

// Use System.IdentityModel.Tokens.Jwt library to validate standard parts
JwtSecurityTokenHandler tokenHandler = new JwtSecurityTokenHandler();
TokenValidationParameters tvp = new TokenValidationParameters();

tvp.ValidateIssuer = false;
tvp.ValidateAudience = true;
tvp.ValidAudience = "{URL to add-in}";
tvp.ValidateIssuerSigningKey = true;
// GetSigningKeys downloads the auth metadata doc and
// returns a List<SecurityKey>
tvp.IssuerSigningKeys = GetSigningKeys(authMetadataUrl);
tvp.ValidateLifetime = true;

try
{
    var claimsPrincipal = tokenHandler.ValidateToken(encodedToken, tvp, out SecurityToken validatedToken);

    // If no exception, all standard checks passed
}
catch (SecurityTokenValidationException ex)
{
    // Validation failed
}
```

<br/>

<span data-ttu-id="a675b-161">La classe `ExchangeAppContext` est définie comme suit :</span><span class="sxs-lookup"><span data-stu-id="a675b-161">The `ExchangeAppContext` class is defined as follows:</span></span>

```cs
using Newtonsoft.Json;

/// <summary>
/// Representation of the appctx claim in an Exchange user identity token.
/// </summary>
public class ExchangeAppContext
{
    /// <summary>
    /// The Exchange identifier for the user
    /// </summary>
    [JsonProperty("msexchuid")]
    public string ExchangeUid { get; set; }

    /// <summary>
    /// The token version
    /// </summary>
    public string Version { get; set; }

    /// <summary>
    /// The URL to download authentication metadata
    /// </summary>
    [JsonProperty("amurl")]
    public string MetadataUrl { get; set; }
}
```

<span data-ttu-id="a675b-162">Pour un exemple qui utilise cette bibliothèque pour valider les jetons Exchange et qui a une implémentation de `GetSigningKeys`, reportez-vous à [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer).</span><span class="sxs-lookup"><span data-stu-id="a675b-162">For an example that uses this library to validate Exchange tokens and has an implementation of `GetSigningKeys`, see [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer).</span></span>

## <a name="see-also"></a><span data-ttu-id="a675b-163">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="a675b-163">See also</span></span>

- [<span data-ttu-id="a675b-164">Outlook-Add-In-Token-Viewer</span><span class="sxs-lookup"><span data-stu-id="a675b-164">Outlook-Add-In-Token-Viewer</span></span>](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)
- [<span data-ttu-id="a675b-165">Outlook-Add-in-JavaScript-ValidateIdentityToken</span><span class="sxs-lookup"><span data-stu-id="a675b-165">Outlook-Add-in-JavaScript-ValidateIdentityToken</span></span>](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ValidateIdentityToken)
