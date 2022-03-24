---
title: Authentifier un utilisateur avec un jeton identité dans un complément
description: Découvrez comment utiliser le jeton d’identité fourni par un complément Outlook pour implémenter l’authentification unique SSO dans votre service.
ms.date: 10/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: 5f4dd8345de0edaaef333ee2b01890e876e049a6
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744620"
---
# <a name="authenticate-a-user-with-an-identity-token-for-exchange"></a>Authentifier un utilisateur avec un jeton d’identité pour Exchange

Les jetons d’identité d’utilisateur Exchange sont un moyen pour vos compléments d'identifier leurs utilisateurs de manière unique.
 En établissant l’identité de l’utilisateur, vous pouvez implémenter un schéma d’authentification unique (SSO) pour votre service back-end qui permet aux clients qui utilisent des modules complémentaires Outlook de se connecter à votre service sans se connecter. Pour plus d’informations sur l’utilisation de ce type de jeton, voir [Jeton d’identité d’utilisateur Exchange](authentication.md#exchange-user-identity-token). Dans cet article, nous allons examiner une méthode simple pour authentifier un utilisateur sur votre back end à l’aide d’un jeton d’identité Exchange.


> [!IMPORTANT]
> Il s’agit tout simplement d’un exemple d’implémentation d’une authentification unique. Comme toujours, lorsqu’il est question d’identité et d’authentification, vous devez vous assurer que votre code respecte les exigences en matière de sécurité de votre organisation.

## <a name="send-the-id-token-with-each-request"></a>Envoyer le jeton d’ID avec chaque requête

La première étape concerne votre complément qui doit obtenir du serveur le jeton d’identité d’utilisateur Exchange en appelant la méthode [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods). Le complément envoie ensuite ce jeton avec chaque requête effectuée à votre serveur principal. Cela peut se faire dans un en-tête ou dans le corps de la requête.

## <a name="validate-the-token"></a>Valider le jeton

Le serveur principal DOIT valider le jeton avant de l’accepter. Il s’agit d’une étape importante pour garantir que le jeton a été émis par le serveur Exchange de l’utilisateur. Pour plus d’informations sur la validation des jetons d’identité d’utilisateur Exchange, reportez-vous à l’article [Valider un jeton d’identité Exchange](validate-an-identity-token.md).

Une fois validée et décodée, la charge utile du jeton ressemble à ce qui suit :

```json
{ 
    "aud" : "https://mailhost.contoso.com/IdentityTest.html",
    "iss" : "00000002-0000-0ff1-ce00-000000000000@mailhost.contoso.com",
    "nbf" : "1505749527",
    "exp" : "1505778327",
    "appctxsender":"00000002-0000-0ff1-ce00-000000000000@mailhost.context.com",
    "isbrowserhostedapp":"true",
    "appctx" : {
        "msexchuid" : "53e925fa-76ba-45e1-be0f-4ef08b59d389",
        "version" : "ExIdTok.V1",
        "amurl" : "https://mailhost.contoso.com:443/autodiscover/metadata/json/1"
    }
}
```

## <a name="map-the-token-to-a-user-in-your-backend"></a>Mapper le jeton à un utilisateur dans votre back end


Votre service principal peut calculer un ID d’utilisateur unique à partir du jeton et le mapper à un utilisateur dans votre système d’utilisateur interne. Par exemple, si vous utilisez une base de données pour stocker des utilisateurs, vous pouvez ajouter cet ID unique à l’enregistrement de l’utilisateur dans votre base de données.

### <a name="generate-a-unique-id"></a>Génération d’un ID unique

Utilisez une combinaison des propriétés `msexchuid` et des `amurl` propriétés. Par exemple, vous pouvez concaténer les deux valeurs et générer une chaîne codée au format base64. Cette valeur peut être générée en toute fiabilité à partir du jeton à chaque fois. Ainsi, vous pouvez mapper un jeton d’identité d’utilisateur Exchange à l’utilisateur dans votre système.

### <a name="check-the-user"></a>Vérification de l’utilisateur

Avec l’ID unique généré, l’étape suivante consiste à vérifier la présence d’un utilisateur dans votre système avec cet ID associé.

- Si vous trouvez l’utilisateur, le back end considère la requête comme authentifiée et autorise sa poursuite.


- Si l’utilisateur est introuvable, le back end renvoie une erreur indiquant que l’utilisateur doit se connecter. 
 Le complément invite ensuite l’utilisateur à se connecter au back end à l’aide de votre méthode d’authentification.
 Une fois l’utilisateur authentifié, le jeton d’identité d’utilisateur Exchange est envoyé avec les détails de l’authentification utilisateur. Le back end peut ensuite mettre à jour l’enregistrement de l’utilisateur dans votre système avec l’ID unique.

