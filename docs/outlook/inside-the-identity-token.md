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
# <a name="inside-the-exchange-identity-token"></a>Présentation du jeton d’identité Exchange

Le jeton d’identité d’utilisateur Exchange renvoyé par la méthode [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) permet au code du complément d’inclure l’identité de l’utilisateur avec des appels à votre service principal. Cet article présente le format et le contenu du jeton.

Un jeton d’identité d’utilisateur Exchange est une chaîne d’URL encodée au format base64 signée par le serveur Exchange qui l’a envoyée. Le jeton n’est pas chiffré et la clé publique qui permet de valider la signature est stockée sur le serveur Exchange qui a émis le jeton. Le jeton comporte trois parties : un en-tête, une charge utile et une signature. Dans la chaîne du jeton, les parties sont séparées par un point (`.`) pour faciliter le fractionnement du jeton.

Exchange utilise le format JSON Web Token (JWT) pour le jeton d’identité. Pour plus d’informations sur les jetons JWT, reportez-vous au document [RFC 7519 JSON Web Token (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt).

## <a name="identity-token-header"></a>En-tête du jeton d’identité

L’en-tête fournit des informations sur le format et la signature du jeton. L’exemple suivant illustre l’en-tête du jeton.

```JSON
{
  "typ": "JWT",
  "alg": "RS256",
  "x5t": "Un6V7lYN-rMgaCoFSTO5z707X-4"
}
```

<br/>
 
Le tableau suivant décrit les parties de l’en-tête du jeton.

| Revendication | Valeur | Description |
|:-----|:-----|:-----|
| `typ` | `JWT` | Identifie le jeton comme un jeton Web JSON. Tous les jetons d’identité fournis par le serveur Exchange sont des jetons JWT. |
| `alg` | `RS256` | L’algorithme de hachage est utilisé pour créer la signature. Tous les jetons fournis par le serveur Exchange utilisent RSASSA-PKCS1-v1_5 avec l’algorithme de hachage SHA-256. |
| `x5t` | Empreinte de certificat | L’empreinte X.509 du jeton. |

## <a name="identity-token-payload"></a>Charge utile du jeton d’identité

La charge utile contient les revendications d’authentification qui identifient le compte de messagerie et identifient le serveur Exchange qui a envoyé le jeton. L’exemple suivant montre à quoi ressemble la section de charge utile.

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
 
Le tableau suivant répertorie les différentes parties de la charge utile du jeton d’identité.

| Revendication | Description |
|:-----|:-----|
| `aud` | L’URL du complément ayant demandé le jeton. Un jeton est valide uniquement s’il est envoyé par le complément en cours d’exécution dans le navigateur du client. Si le complément utilise la version 1.1 du schéma des manifestes des compléments Office, cette URL correspond à celle indiquée dans le premier élément `SourceLocation`, sous le type de formulaire `ItemRead` ou `ItemEdit`, selon celui qui apparaît en premier dans l’élément [FormSettings](../reference/manifest/formsettings.md) du manifeste de complément. |
| `iss` | Un identificateur unique du serveur Exchange qui a émis le jeton. Tous les jetons émis par ce serveur Exchange auront le même identificateur. |
| `nbf` | La date et l’heure de début de validité du jeton. La valeur correspond au nombre de secondes depuis le 1er janvier 1970. |
| `exp` | La date et l’heure de fin de validité du jeton. La valeur correspond au nombre de secondes depuis le 1er janvier 1970. |
| `appctxsender` | Identificateur unique du serveur Exchange qui a envoyé le contexte de l’application. |
| `isbrowserhostedapp` | Indique si le complément est hébergé dans un navigateur. |
| `appctx` | Contexte d’application du jeton. |

Les informations contenues dans la réclamation appctx fournissent l’identificateur unique pour le compte et l’emplacement de la clé publique utilisée pour signer le jeton. Le tableau suivant répertorie les parties de la réclamation `appctx`.

| Propriété du contexte de l’application | Description |
|:-----|:-----|
| `msexchuid` | Identificateur unique associé au compte de messagerie et au serveur Exchange. |
| `version` | Numéro de version du jeton. Pour tous les jetons fournis par Exchange, la valeur est `ExIdTok.V1`. |
| `amurl` | URL du document de métadonnées d’authentification qui contient la clé publique du certificat X.509 utilisé pour signer le jeton.<br/><br/>Pour plus d’informations sur l’utilisation du document de métadonnées d’authentification, reportez-vous à [Valider un jeton d’identité Exchange](validate-an-identity-token.md). |

## <a name="identity-token-signature"></a>Signature du jeton d’identité

La signature est créée par hachage des sections d’en-tête et de charge utile avec l’algorithme spécifié dans l’en-tête et en utilisant le certificat X509 autosigné situé sur le serveur à l’emplacement spécifié dans la charge utile. Votre service web peut valider cette signature pour contribuer à assurer que le jeton d’identité provient bien du serveur prévu pour son envoie.

## <a name="see-also"></a>Voir aussi

Pour consulter un exemple d’analyse du jeton d’identité d’utilisateur Exchange, reportez-vous à [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer).
