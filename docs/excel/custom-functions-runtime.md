---
ms.date: 09/25/2020
description: Comprendre Excel fonctions personnalisées qui n’utilisent pas de volet de tâches et leur runtime JavaScript spécifique.
title: Runtime pour les fonctions personnalisées sans interface Excel’interface utilisateur
localization_priority: Normal
ms.openlocfilehash: aa2cf2632ddf9eb1ad1eb202b031ee2ca686af01
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349622"
---
# <a name="runtime-for-ui-less-excel-custom-functions"></a>Runtime pour les fonctions personnalisées sans interface Excel’interface utilisateur

Les fonctions personnalisées qui n’utilisent pas de volet de tâches (fonctions personnalisées sans interface utilisateur) utilisent un runtime JavaScript conçu pour optimiser les performances des calculs.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

Ce runtime JavaScript permet d’accéder aux API de l’espace de noms qui peuvent être utilisées par les fonctions personnalisées sans interface utilisateur et le volet Des tâches pour stocker `OfficeRuntime` des données.

## <a name="requesting-external-data"></a>Demande de données externes

Dans une fonction personnalisée sans interface utilisateur, vous pouvez demander des données externes à l’aide d’une API telle que [Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) ou en utilisant [XmlHttpRequest (XHR),](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest)une API web standard qui émettre des demandes HTTP pour interagir avec les serveurs.

N’ignorez pas que les fonctions sans interface utilisateur doivent utiliser des mesures [](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy) de sécurité supplémentaires lors de la génération de XmlHttpRequests, nécessitant une stratégie d’origine identique et [un CORS](https://www.w3.org/TR/cors/)simple.

Une implémentation CORS simple ne peut pas utiliser de cookies et prend uniquement en charge les méthodes simples (GET, HEAD, POST). Le simple CORS accepte des en-têtes simples avec des noms de champs `Accept`, `Accept-Language`, `Content-Language`. Vous pouvez également utiliser un `Content-Type` en-tête dans CORS simple, à condition que le type de contenu `application/x-www-form-urlencoded` soit , ou `text/plain` `multipart/form-data` .

## <a name="storing-and-accessing-data"></a>Accès aux données et stockage

Dans une fonction personnalisée sans interface utilisateur, vous pouvez stocker et accéder aux données à l’aide de `OfficeRuntime.storage` l’objet. `Storage` est un système de stockage persistant, non chiffré et à valeur clé qui fournit une alternative à [localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage), qui ne peut pas être utilisé par des fonctions personnalisées sans interface utilisateur. `Storage` offre 10 Mo de données par domaine. Les domaines peuvent être partagés par plusieurs modules.

`Storage` est conçu comme une solution de stockage partagé, ce qui signifie que plusieurs parties d’un complément ont accès aux mêmes données. Par exemple, les jetons pour l’authentification des utilisateurs peuvent être stockés, car ils sont accessibles à la fois par une fonction personnalisée sans interface utilisateur et par des éléments d’interface utilisateur de add-in tels qu’un volet Des `storage` tâches. De même, si deux modules complémentaires partagent le même domaine (par exemple, , ), ils sont également autorisés à partager des informations entre `www.contoso.com/addin1` `www.contoso.com/addin2` `storage` eux. Notez que les add-ins qui ont différents sous-domaine auront différentes instances `storage` de (par exemple, `subdomain.contoso.com/addin1` , `differentsubdomain.contoso.com/addin2` ).

Comme `storage` peut être un emplacement partagé, il est important de savoir qu’il est possible de remplacer des paires clé-valeur.

Les méthodes suivantes sont disponibles sur `storage` l’objet.

- `getItem`
- `getItems`
- `setItem`
- `setItems`
- `removeItem`
- `removeItems`
- `getKeys`

> [!NOTE]
> Il n’existe aucune méthode pour effacer toutes les informations (par `clear` exemple). À la place, vous devez utiliser l’objet `removeItems` pour supprimer plusieurs entrées à la fois.

### <a name="officeruntimestorage-example"></a>Exemple OfficeRuntime.storage

L’exemple de code suivant appelle `OfficeRuntime.storage.setItem` la fonction pour définir une clé et une valeur dans `storage` .

```js
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## <a name="additional-considerations"></a>Considérations supplémentaires

Si votre add-in utilise uniquement des fonctions personnalisées sans interface utilisateur, notez que vous ne pouvez pas accéder au modèle DOM (Document Object Model) avec des fonctions personnalisées sans interface utilisateur ou utiliser des bibliothèques telles que jQuery qui reposent sur le DOM.

## <a name="next-steps"></a>Étapes suivantes
Découvrez comment [déboguer des](custom-functions-debugging.md)fonctions personnalisées sans interface utilisateur.

## <a name="see-also"></a>Voir aussi

* [Authentifier les fonctions personnalisées sans interface utilisateur](custom-functions-authentication.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
* [Didacticiel sur les fonctions personnalisées](../tutorials/excel-tutorial-create-custom-functions.md)
