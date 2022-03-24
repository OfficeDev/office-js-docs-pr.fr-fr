---
title: Activer l’sign-on unique (SSO) dans Outlook compléments qui utilisent l’activation basée sur des événements
description: Découvrez comment activer l’utilisateur unique lorsque vous travaillez dans un complément d’activation basé sur des événements.
ms.date: 03/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 38c717e0d626f4c135f76350e30398db26cac24f
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746539"
---
# <a name="enable-single-sign-on-sso-in-outlook-add-ins-that-use-event-based-activation"></a>Activer l’sign-on unique (SSO) dans Outlook compléments qui utilisent l’activation basée sur des événements

Lorsqu’Outlook complément utilise l’activation basée sur des événements, les événements s’exécutent dans un runtime JavaScript distinct. Après avoir effectué les étapes de l’authentification d’un utilisateur avec un jeton d’authentification unique dans un complément [Outlook](authenticate-a-user-with-an-sso-token.md), suivez les étapes supplémentaires décrites dans cet article pour activer l’authentification unique pour votre code de gestion des événements. Une fois l’utilisateur unique activé, vous pouvez appeler l’API `getAccessToken()` pour obtenir un jeton d’accès avec l’identité de l’utilisateur.

> [!NOTE]
> Les étapes de cet article s’appliquent uniquement lorsque vous exécutez votre Outlook sur Windows. En effet, Outlook sur Windows utilise un fichier JavaScript, tandis que Outlook sur le web utilise un fichier HTML qui peut référencer le même fichier JavaScript.

Par Outlook sur Windows, dans le manifeste de votre complément Outlook, vous identifiez un seul fichier JavaScript à charger pour l’activation basée sur des événements. Vous devez également spécifier Office que ce fichier est autorisé à prendre en charge l’sso. Pour ce faire, vous créez une liste de tous les modules et de leurs fichiers JavaScript pour les Office via un URI connu.

## <a name="list-allowed-add-ins-with-a-well-known-uri"></a>Liste des add-ins autorisés avec un URI connu

Pour ré lister les modules qui sont autorisés à fonctionner avec l’chacune d’elles, créez un fichier JSON qui identifie chaque fichier JavaScript pour chaque module. Hébergez ensuite ce fichier JSON à un URI connu. Un URI connu permet la spécification de tous les fichiers JS hébergés autorisés à obtenir des jetons pour l’origine web actuelle. Cela garantit que le propriétaire de l’origine dispose d’un contrôle total sur les fichiers JS hébergés qui sont destinés à être utilisés dans un add-in et ceux qui ne le sont pas, ce qui empêche toute faille de sécurité autour de l’emprunt d’identité, par exemple.

L’exemple suivant montre comment activer l’utilisateur principal pour deux modules (version principale et version bêta). Vous pouvez lister autant de modules que nécessaire en fonction du nombre que vous fournissez à partir de votre serveur web.

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

L’origine fait référence à un modèle de schéma + sous-domaine + domaine + port. Le nom de l’emplacement **doit** être `.well-known`et le nom du fichier de ressources **doit être** `microsoft-officeaddins-allowed.json`. Ce fichier doit contenir un objet JSON `allowed` avec un attribut nommé dont la valeur est un tableau de tous les fichiers JavaScript autorisés pour l’sso pour leurs add-ins respectifs.

## <a name="see-also"></a>Voir aussi

- [Authentifier un utilisateur avec un jeton d’authentification unique dans un Outlook d’authentification unique](authenticate-a-user-with-an-sso-token.md)
- [Configurer votre complément Outlook pour l’activation basée sur des événements](autolaunch.md)
