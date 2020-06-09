---
title: Présentation du jeton d’identité Exchange dans un complément Outlook
description: Découvrez le contenu d’un jeton d’identité d’utilisateur Exchange généré à partir d’un complément Outlook.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: dee8416660386c25a55caa42b6e5ee8685ee8852
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609089"
---
# <a name="inside-the-exchange-identity-token"></a><span data-ttu-id="3a939-103">Présentation du jeton d’identité Exchange</span><span class="sxs-lookup"><span data-stu-id="3a939-103">Inside the Exchange identity token</span></span>

<span data-ttu-id="3a939-104">Le jeton d’identité d’utilisateur Exchange renvoyé par la méthode [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) permet au code du complément d’inclure l’identité de l’utilisateur avec des appels à votre service principal.</span><span class="sxs-lookup"><span data-stu-id="3a939-104">The Exchange user identity token returned by the [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method provides a way for your add-in code to include the user's identity with calls to your back-end service.</span></span> <span data-ttu-id="3a939-105">Cet article présente le format et le contenu du jeton.</span><span class="sxs-lookup"><span data-stu-id="3a939-105">This article will discuss the format and contents of the token.</span></span>

<span data-ttu-id="3a939-106">Un jeton d’identité d’utilisateur Exchange est une chaîne d’URL encodée au format base64 signée par le serveur Exchange qui l’a envoyée.</span><span class="sxs-lookup"><span data-stu-id="3a939-106">An Exchange user identity token is a base-64 URL-encoded string that is signed by the Exchange server that sent it.</span></span> <span data-ttu-id="3a939-107">Le jeton n’est pas chiffré et la clé publique qui permet de valider la signature est stockée sur le serveur Exchange qui a émis le jeton.</span><span class="sxs-lookup"><span data-stu-id="3a939-107">The token is not encrypted, and the public key that you use to validate the signature is stored on the Exchange server that issued the token.</span></span> <span data-ttu-id="3a939-108">Le jeton comporte trois parties : un en-tête, une charge utile et une signature.</span><span class="sxs-lookup"><span data-stu-id="3a939-108">The token has three parts: a header, a payload, and a signature.</span></span> <span data-ttu-id="3a939-109">Dans la chaîne du jeton, les parties sont séparées par un point (`.`) pour faciliter le fractionnement du jeton.</span><span class="sxs-lookup"><span data-stu-id="3a939-109">In the token string, the parts are separated by a period character (`.`) to make it easy for you to split the token.</span></span>

<span data-ttu-id="3a939-110">Exchange utilise le format JSON Web Token (JWT) pour le jeton d’identité.</span><span class="sxs-lookup"><span data-stu-id="3a939-110">Exchange uses a the JSON Web Token (JWT) format for the identity token.</span></span> <span data-ttu-id="3a939-111">Pour plus d’informations sur les jetons JWT, reportez-vous au document [RFC 7519 JSON Web Token (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt).</span><span class="sxs-lookup"><span data-stu-id="3a939-111">For information about JWT tokens, see [RFC 7519 JSON Web Token (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt).</span></span>

## <a name="identity-token-header"></a><span data-ttu-id="3a939-112">En-tête du jeton d’identité</span><span class="sxs-lookup"><span data-stu-id="3a939-112">Identity token header</span></span>

<span data-ttu-id="3a939-113">L’en-tête fournit des informations sur le format et la signature du jeton.</span><span class="sxs-lookup"><span data-stu-id="3a939-113">The header provides information about the format and signature information of the token.</span></span> <span data-ttu-id="3a939-114">L’exemple suivant illustre l’en-tête du jeton.</span><span class="sxs-lookup"><span data-stu-id="3a939-114">The following example shows what the header of the token looks like.</span></span>

```JSON
{
  "typ": "JWT",
  "alg": "RS256",
  "x5t": "Un6V7lYN-rMgaCoFSTO5z707X-4"
}
```

<br/>
 
<span data-ttu-id="3a939-115">Le tableau suivant décrit les parties de l’en-tête du jeton.</span><span class="sxs-lookup"><span data-stu-id="3a939-115">The following table describes the parts of the token header.</span></span>

| <span data-ttu-id="3a939-116">Revendication</span><span class="sxs-lookup"><span data-stu-id="3a939-116">Claim</span></span> | <span data-ttu-id="3a939-117">Valeur</span><span class="sxs-lookup"><span data-stu-id="3a939-117">Value</span></span> | <span data-ttu-id="3a939-118">Description</span><span class="sxs-lookup"><span data-stu-id="3a939-118">Description</span></span> |
|:-----|:-----|:-----|
| `typ` | `JWT` | <span data-ttu-id="3a939-119">Identifie le jeton comme un jeton Web JSON.</span><span class="sxs-lookup"><span data-stu-id="3a939-119">Identifies the token as a JSON Web Token.</span></span> <span data-ttu-id="3a939-120">Tous les jetons d’identité fournis par le serveur Exchange sont des jetons JWT.</span><span class="sxs-lookup"><span data-stu-id="3a939-120">All identity tokens provided by Exchange server are JWT tokens.</span></span> |
| `alg` | `RS256` | <span data-ttu-id="3a939-121">L’algorithme de hachage est utilisé pour créer la signature.</span><span class="sxs-lookup"><span data-stu-id="3a939-121">The hashing algorithm that is used to create the signature.</span></span> <span data-ttu-id="3a939-122">Tous les jetons fournis par le serveur Exchange utilisent RSASSA-PKCS1-v1_5 avec l’algorithme de hachage SHA-256.</span><span class="sxs-lookup"><span data-stu-id="3a939-122">All tokens provided by Exchange server use the RSASSA-PKCS1-v1_5 with SHA-256 hash algorithm.</span></span> |
| `x5t` | <span data-ttu-id="3a939-123">Empreinte de certificat</span><span class="sxs-lookup"><span data-stu-id="3a939-123">Certificate thumbprint</span></span> | <span data-ttu-id="3a939-124">L’empreinte X.509 du jeton.</span><span class="sxs-lookup"><span data-stu-id="3a939-124">The X.509 thumbprint of the token.</span></span> |

## <a name="identity-token-payload"></a><span data-ttu-id="3a939-125">Charge utile du jeton d’identité</span><span class="sxs-lookup"><span data-stu-id="3a939-125">Identity token payload</span></span>

<span data-ttu-id="3a939-p107">La charge utile contient les revendications d’authentification qui identifient le compte de messagerie et identifient le serveur Exchange qui a envoyé le jeton. L’exemple suivant montre à quoi ressemble la section de charge utile.</span><span class="sxs-lookup"><span data-stu-id="3a939-p107">The payload contains the authentication claims that identify the email account and identify the Exchange server that sent the token. The following example shows what the payload section looks like.</span></span>

```JSON
{ 
  "aud": "https://mailhost.contoso.com/IdentityTest.html", 
  "iss": "00000002-0000-0ff1-ce00-000000000000@mailhost.contoso.com", 
  "nbf": "1331579055", 
  "exp": "1331607855", 
  "appctxsender": "00000002-0000-0ff1-ce00-000000000000@mailhost.context.com",
  "isbrowserhostedapp": "true",
  "appctx": { 
    "msexchuid": "53e925fa-76ba-45e1-be0f-4ef08b59d389@mailhost.contoso.com",
    "version": "ExIdTok.V1",
    "amurl": "https://mailhost.contoso.com:443/autodiscover/metadata/json/1"
  } 
}
```

<br/>
 
<span data-ttu-id="3a939-128">Le tableau suivant répertorie les différentes parties de la charge utile du jeton d’identité.</span><span class="sxs-lookup"><span data-stu-id="3a939-128">The following table lists the parts of the identity token payload.</span></span>

| <span data-ttu-id="3a939-129">Revendication</span><span class="sxs-lookup"><span data-stu-id="3a939-129">Claim</span></span> | <span data-ttu-id="3a939-130">Description</span><span class="sxs-lookup"><span data-stu-id="3a939-130">Description</span></span> |
|:-----|:-----|
| `aud` | <span data-ttu-id="3a939-131">L’URL du complément ayant demandé le jeton.</span><span class="sxs-lookup"><span data-stu-id="3a939-131">The URL of the add-in that requested the token.</span></span> <span data-ttu-id="3a939-132">Un jeton est valide uniquement s’il est envoyé par le complément en cours d’exécution dans le navigateur du client.</span><span class="sxs-lookup"><span data-stu-id="3a939-132">A token is only valid if it is sent from the add-in that is running in the client's browser.</span></span> <span data-ttu-id="3a939-133">Si le complément utilise la version 1.1 du schéma des manifestes des compléments Office, cette URL correspond à celle indiquée dans le premier élément `SourceLocation`, sous le type de formulaire `ItemRead` ou `ItemEdit`, selon celui qui apparaît en premier dans l’élément [FormSettings](../reference/manifest/formsettings.md) du manifeste de complément.</span><span class="sxs-lookup"><span data-stu-id="3a939-133">If the add-in uses the Office Add-ins manifests schema v1.1, this URL is the URL specified in the first `SourceLocation` element, under the form type `ItemRead` or `ItemEdit`, whichever occurs first as part of the [FormSettings](../reference/manifest/formsettings.md) element in the add-in manifest.</span></span> |
| `iss` | <span data-ttu-id="3a939-p109">Un identificateur unique du serveur Exchange qui a émis le jeton. Tous les jetons émis par ce serveur Exchange auront le même identificateur.</span><span class="sxs-lookup"><span data-stu-id="3a939-p109">A unique identifier for the Exchange server that issued the token. All tokens issued by this Exchange server will have the same identifier.</span></span> |
| `nbf` | <span data-ttu-id="3a939-p110">La date et l’heure de début de validité du jeton. La valeur correspond au nombre de secondes depuis le 1er janvier 1970.</span><span class="sxs-lookup"><span data-stu-id="3a939-p110">The date and time that the token is valid starting from. The value is the number of seconds since January 1, 1970.</span></span> |
| `exp` | <span data-ttu-id="3a939-p111">La date et l’heure de fin de validité du jeton. La valeur correspond au nombre de secondes depuis le 1er janvier 1970.</span><span class="sxs-lookup"><span data-stu-id="3a939-p111">The date and time that the token is valid until. The value is the number of seconds since January 1, 1970.</span></span> |
| `appctxsender` | <span data-ttu-id="3a939-140">Identificateur unique du serveur Exchange qui a envoyé le contexte de l’application.</span><span class="sxs-lookup"><span data-stu-id="3a939-140">A unique identifier for the Exchange server that sent the application context.</span></span> |
| `isbrowserhostedapp` | <span data-ttu-id="3a939-141">Indique si le complément est hébergé dans un navigateur.</span><span class="sxs-lookup"><span data-stu-id="3a939-141">Indicates whether the add-in is hosted in a browser.</span></span> |
| `appctx` | <span data-ttu-id="3a939-142">Contexte d’application du jeton.</span><span class="sxs-lookup"><span data-stu-id="3a939-142">The application context for the token.</span></span> |

<span data-ttu-id="3a939-143">Les informations contenues dans la réclamation appctx fournissent l’identificateur unique pour le compte et l’emplacement de la clé publique utilisée pour signer le jeton.</span><span class="sxs-lookup"><span data-stu-id="3a939-143">The information in the appctx claim provides you with the unique identifier for the account and the location of the public key used to sign the token.</span></span> <span data-ttu-id="3a939-144">Le tableau suivant répertorie les parties de la réclamation `appctx`.</span><span class="sxs-lookup"><span data-stu-id="3a939-144">The following table lists the parts of the `appctx` claim.</span></span>

| <span data-ttu-id="3a939-145">Propriété du contexte de l’application</span><span class="sxs-lookup"><span data-stu-id="3a939-145">Application context property</span></span> | <span data-ttu-id="3a939-146">Description</span><span class="sxs-lookup"><span data-stu-id="3a939-146">Description</span></span> |
|:-----|:-----|
| `msexchuid` | <span data-ttu-id="3a939-147">Identificateur unique associé au compte de messagerie et au serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="3a939-147">A unique identifier associated with the email account and the Exchange server.</span></span> |
| `version` | <span data-ttu-id="3a939-148">Numéro de version du jeton.</span><span class="sxs-lookup"><span data-stu-id="3a939-148">The version number of the token.</span></span> <span data-ttu-id="3a939-149">Pour tous les jetons fournis par Exchange, la valeur est `ExIdTok.V1`.</span><span class="sxs-lookup"><span data-stu-id="3a939-149">For all tokens provided by Exchange, the value is `ExIdTok.V1`.</span></span> |
| `amurl` | <span data-ttu-id="3a939-150">URL du document de métadonnées d’authentification qui contient la clé publique du certificat X.509 utilisé pour signer le jeton.</span><span class="sxs-lookup"><span data-stu-id="3a939-150">The URL of the authentication metadata document that contains the public key of the X.509 certificate that was used to sign the token.</span></span><br/><br/><span data-ttu-id="3a939-151">Pour plus d’informations sur l’utilisation du document de métadonnées d’authentification, reportez-vous à [Valider un jeton d’identité Exchange](validate-an-identity-token.md).</span><span class="sxs-lookup"><span data-stu-id="3a939-151">For more information about how to use the authentication metadata document, see [Validate an Exchange identity token](validate-an-identity-token.md).</span></span> |

## <a name="identity-token-signature"></a><span data-ttu-id="3a939-152">Signature du jeton d’identité</span><span class="sxs-lookup"><span data-stu-id="3a939-152">Identity token signature</span></span>

<span data-ttu-id="3a939-p114">La signature est créée par hachage des sections d’en-tête et de charge utile avec l’algorithme spécifié dans l’en-tête et en utilisant le certificat X509 autosigné situé sur le serveur à l’emplacement spécifié dans la charge utile. Votre service web peut valider cette signature pour contribuer à assurer que le jeton d’identité provient bien du serveur prévu pour son envoie.</span><span class="sxs-lookup"><span data-stu-id="3a939-p114">The signature is created by hashing the header and payload sections with the algorithm specified in the header and using the self-signed X509 certificate located on the server at the location specified in the payload. Your web service can validate this signature to help make sure that the identity token comes from the server that you expect to send it.</span></span>

## <a name="see-also"></a><span data-ttu-id="3a939-155">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="3a939-155">See also</span></span>

<span data-ttu-id="3a939-156">Pour consulter un exemple d’analyse du jeton d’identité d’utilisateur Exchange, reportez-vous à [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer).</span><span class="sxs-lookup"><span data-stu-id="3a939-156">For an example that parses the Exchange user identity token, see [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer).</span></span>
