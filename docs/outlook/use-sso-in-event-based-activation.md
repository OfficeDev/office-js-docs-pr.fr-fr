---
title: Activer l’authentification unique (SSO) dans les compléments Outlook qui utilisent l’activation basée sur les événements
description: Découvrez comment activer l’authentification unique lors de l’utilisation d’un complément d’activation basé sur des événements.
ms.date: 06/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 10fd973c0476878443d7238e8805aa4db9f62953
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423117"
---
# <a name="enable-single-sign-on-sso-in-outlook-add-ins-that-use-event-based-activation"></a>Activer l’authentification unique (SSO) dans les compléments Outlook qui utilisent l’activation basée sur les événements

Lorsqu’un complément Outlook utilise l’activation basée sur les événements, les événements s’exécutent dans un [runtime](../testing/runtimes.md) distinct. Après avoir effectué les étapes décrites dans [Authentifier un utilisateur avec un jeton d’authentification unique dans un complément Outlook](authenticate-a-user-with-an-sso-token.md), suivez les étapes supplémentaires décrites dans cet article pour activer l’authentification unique pour votre code de gestion des événements. Une fois que vous avez activé l’authentification unique, vous pouvez appeler [l’API getAccessToken()](/javascript/api/office-runtime/officeruntime.auth) pour obtenir un jeton d’accès avec l’identité de l’utilisateur.

> [!IMPORTANT]
> `Office.auth.getAccessToken` Tout en `OfficeRuntime.auth.getAccessToken` effectuant les mêmes fonctionnalités de récupération d’un jeton d’accès, nous vous recommandons d’appeler `OfficeRuntime.auth.getAccessToken` votre complément basé sur les événements. Cette API est prise en charge dans toutes les versions du client Outlook qui prennent en charge l’activation basée sur les événements et l’authentification unique. En revanche, `Office.auth.getAccessToken` est uniquement pris en charge dans Outlook sur Windows à partir de la version 2111 (build 14701.20000).

Pour Outlook sur Windows, dans le manifeste de votre complément Outlook, vous identifiez un seul fichier JavaScript à charger pour l’activation basée sur les événements. Vous devez également spécifier à Office que ce fichier est autorisé à prendre en charge l’authentification unique. Pour ce faire, vous créez une liste de tous les compléments, ainsi que leurs fichiers JavaScript, à fournir à Office via un URI connu.

> [!NOTE]
> Les étapes décrites dans cet article s’appliquent uniquement lors de l’exécution de votre complément Outlook sur Windows. Cela est dû au fait qu’Outlook sur Windows utilise un fichier JavaScript, tandis que Outlook sur le web utilise un fichier HTML qui peut référencer le même fichier JavaScript.

## <a name="list-allowed-add-ins-with-a-well-known-uri"></a>Répertorier les compléments autorisés avec un URI connu

Pour répertorier les compléments autorisés à fonctionner avec l’authentification unique, créez un fichier JSON qui identifie chaque fichier JavaScript pour chaque complément. Ensuite, hébergez ce fichier JSON à un URI connu. Un URI connu permet la spécification de tous les fichiers JS hébergés autorisés à obtenir des jetons pour l’origine web actuelle. Cela garantit que le propriétaire de l’origine dispose d’un contrôle total sur les fichiers JS hébergés destinés à être utilisés dans un complément et sur les fichiers qui ne le sont pas, ce qui empêche les failles de sécurité liées à l’emprunt d’identité, par exemple.

L’exemple suivant montre comment activer l’authentification unique pour deux compléments (une version principale et une version bêta). Vous pouvez répertorier autant de compléments que nécessaire en fonction du nombre que vous fournissez à partir de votre serveur web.

```json
{
    "allowed":
    [
        "https://addin.contoso.com:8000/main/js/autorun.js",
        "https://addin.contoso.com:8000/beta/js/autorun.js"
    ]
}
```

Hébergez le fichier JSON sous un emplacement nommé `.well-known` dans l’URI à la racine de l’origine. Par exemple, si l’origine est `https://addin.contoso.com:8000/`, l’URI connu est `https://addin.contoso.com:8000/.well-known/microsoft-officeaddins-allowed.json`.

L’origine fait référence à un modèle de schéma + sous-domaine + domaine + port. Le nom de l’emplacement **doit** être `.well-known`, et le nom du fichier de ressources **doit** être `microsoft-officeaddins-allowed.json`. Ce fichier doit contenir un objet JSON avec un attribut nommé `allowed` dont la valeur est un tableau de tous les fichiers JavaScript autorisés pour l’authentification unique pour leurs compléments respectifs.

## <a name="see-also"></a>Voir aussi

- [Authentifier un utilisateur avec un jeton d’authentification unique dans un complément Outlook](authenticate-a-user-with-an-sso-token.md)
- [Configurer votre complément Outlook pour l’activation basée sur les événements](autolaunch.md)
