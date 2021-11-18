---
title: Activer l' sign-on unique (SSO) dans Outlook compléments qui utilisent l’activation basée sur des événements
description: Découvrez comment activer l' utilisateur unique lorsque vous travaillez dans un complément d’activation basé sur des événements.
ms.date: 11/16/2021
ms.localizationpriority: medium
ms.openlocfilehash: 66d1edb8b7b0092ee107b73af24d5420caee8677
ms.sourcegitcommit: 6e6c4803fdc0a3cc2c1bcd275288485a987551ff
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/18/2021
ms.locfileid: "61066652"
---
# <a name="enable-single-sign-on-sso-in-outlook-add-ins-that-use-event-based-activation"></a>Activer l' sign-on unique (SSO) dans Outlook compléments qui utilisent l’activation basée sur des événements

Lorsqu’Outlook complément utilise l’activation basée sur des événements, les événements s’exécutent dans un runtime JavaScript distinct. Après avoir effectué les étapes de l’authentification d’un utilisateur avec un jeton d’authentification unique dans un complément [Outlook,](authenticate-a-user-with-an-sso-token.md)suivez les étapes supplémentaires décrites dans cet article pour activer l’authentification unique pour votre code de gestion des événements. Une fois l' utilisateur unique activé, vous pouvez appeler l’API pour obtenir un jeton `getAccessToken()` d’accès avec l’identité de l’utilisateur.

> [!NOTE]
> Les étapes de cet article s’appliquent uniquement lorsque vous exécutez votre Outlook sur Windows. En effet, Outlook sur Windows utilise un fichier JavaScript, tandis que Outlook sur le web utilise un fichier HTML qui peut référencer le même fichier JavaScript.

Par Outlook sur Windows, dans le manifeste de votre complément Outlook, vous identifiez un seul fichier JavaScript à charger pour l’activation basée sur des événements. Vous devez également spécifier Office que ce fichier est autorisé à prendre en charge l' sso. Il existe deux approches à cette fin. Vous pouvez créer une liste de tous les modules et de leurs fichiers JavaScript pour les Office via un URI connu. Vous pouvez également ajouter un en-tête de réponse personnalisé pour activer l’personnalisation de l’personnalisation.

## <a name="list-allowed-add-ins-with-a-well-known-uri"></a>Liste des add-ins autorisés avec un URI connu

Pour ré lister les modules qui sont autorisés à fonctionner avec l' chacune d’elles, créez un fichier JSON qui identifie chaque fichier JavaScript pour chaque module. Hébergez ensuite ce fichier JSON à un URI connu. Un URI connu permet la spécification de tous les fichiers JS hébergés autorisés à obtenir des jetons pour l’origine web actuelle. Cela garantit que le propriétaire de l’origine dispose d’un contrôle total sur les fichiers JS hébergés qui sont destinés à être utilisés dans un add-in et ceux qui ne le sont pas, ce qui empêche toute faille de sécurité autour de l’emprunt d’identité, par exemple.

L’exemple suivant montre comment activer l' utilisateur principal pour deux modules (version principale et version bêta). Vous pouvez lister autant de modules que nécessaire en fonction du nombre que vous fournissez à partir de votre serveur web.

```json
{
    "allowed":
    [
        "https://addin.contoso.com:8000/main/js/autorun.js",
        "https://addin.contoso.com:8000/beta/js/autorun.js"
    ]
}
```

Hébergez le fichier JSON sous un emplacement nommé `.well-known` dans l’URI à la racine de l’origine. Par exemple, si l’origine est `https://addin.contoso.com:8000/` , l’URI connu est `https://addin.contoso.com:8000/.well-known/microsoft-officeaddins-allowed.json` .

L’origine fait référence à un modèle de schéma + sous-domaine + domaine + port. Le nom de l’emplacement **doit** être `.well-known` et le nom du fichier de ressources doit **être** `microsoft-officeaddins-allowed.json` . Ce fichier doit contenir un objet JSON avec un attribut nommé dont la valeur est un tableau de tous les fichiers JavaScript autorisés pour l' sso pour leurs `allowed` add-ins respectifs.

## <a name="add-a-custom-response-header"></a>Ajouter un en-tête de réponse personnalisé

Une deuxième approche consiste à ajouter un en-tête de réponse personnalisé nommé `MS-OfficeAddins-Allowed-Origin` . La valeur de l’en-tête doit être l’origine du fichier JavaScript.

Par exemple, si le fichier JavaScript se trouve à l’emplacement , ajoutez l’en-tête `https://addin.contoso.com:8000/main/js/autorun.js` de réponse suivant.

`MS-OfficeAddins-Allowed-Origin : https://addin.contoso.com:8000`

Vous devez consulter la documentation de votre serveur web spécifique pour savoir comment ajouter l’en-tête de réponse personnalisée.

## <a name="see-also"></a>Voir aussi

- [Authentifier un utilisateur avec un jeton d’authentification unique dans un Outlook’authentification unique](authenticate-a-user-with-an-sso-token.md)
- [Configurer votre complément Outlook pour l’activation basée sur des événements](autolaunch.md)
