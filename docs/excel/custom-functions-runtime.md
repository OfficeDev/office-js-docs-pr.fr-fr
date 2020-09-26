---
ms.date: 09/25/2020
description: Découvrez les fonctions personnalisées Excel qui n’utilisent pas de volet de tâches ni leur propre Runtime JavaScript.
title: Runtime pour les fonctions personnalisées Excel sans interface utilisateur
localization_priority: Normal
ms.openlocfilehash: 94254dfb5a0d03b7c9fec392b2377aff91b58af4
ms.sourcegitcommit: b47318a24a50443b0579e05e178b3bb5433c372f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/25/2020
ms.locfileid: "48279508"
---
# <a name="runtime-for-ui-less-excel-custom-functions"></a>Runtime pour les fonctions personnalisées Excel sans interface utilisateur

Les fonctions personnalisées qui n’utilisent pas de volet de tâches (fonctions personnalisées sans interface utilisateur) utilisent un Runtime JavaScript conçu pour optimiser les performances des calculs.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

Ce Runtime JavaScript fournit l’accès aux API dans l' `OfficeRuntime` espace de noms qui peut être utilisé par les fonctions personnalisées sans interface utilisateur et le volet de tâches pour le stockage des données.

## <a name="requesting-external-data"></a>Demande de données externes

Au sein d’une fonction personnalisée sans interface utilisateur, vous pouvez demander des données externes à l’aide d’une API telle que [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) ou à l’aide de [XMLHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), une API Web standard qui émet des requêtes http pour interagir avec les serveurs.

N’oubliez pas que les fonctions sans interface utilisateur doivent utiliser des mesures de sécurité supplémentaires lors de la création de XmlHttpRequest, nécessitant la [même stratégie d’origine](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) et la même [cors](https://www.w3.org/TR/cors/)simple.

Une implémentation CORS simple ne peut pas utiliser les cookies et ne prend en charge que des méthodes simples (GET, HEAD, POST). Le simple CORS accepte des en-têtes simples avec des noms de champs `Accept`, `Accept-Language`, `Content-Language`. Vous pouvez également utiliser un `Content-Type` en-tête dans un simple cors, à condition que le type de contenu soit `application/x-www-form-urlencoded` , `text/plain` , ou `multipart/form-data` .

## <a name="storing-and-accessing-data"></a>Accès aux données et stockage

Au sein d’une fonction personnalisée sans interface utilisateur, vous pouvez stocker et accéder aux données à l’aide de l' `OfficeRuntime.storage` objet. `Storage` est un système de stockage de valeur de clé persistante, non chiffré qui fournit une alternative à [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), qui ne peut pas être utilisé par des fonctions personnalisées sans interface utilisateur. `Storage` offre 10 Mo de données par domaine. Les domaines peuvent être partagés par plusieurs compléments.

`Storage` est conçu comme une solution de stockage partagé, ce qui signifie que plusieurs parties d’un complément ont accès aux mêmes données. Par exemple, les jetons pour l’authentification utilisateur peuvent être stockés dans `storage` , car il est accessible à la fois par une fonction personnalisée sans interface utilisateur et par des éléments d’interface utilisateur de complément tels qu’un volet de tâches. De même, si deux compléments partagent le même domaine (par exemple, `www.contoso.com/addin1` `www.contoso.com/addin2` ), ils sont également autorisés à partager des informations entre eux `storage` . Notez que les compléments qui ont des sous-domaines différents auront des instances différentes de `storage` (par exemple `subdomain.contoso.com/addin1` , `differentsubdomain.contoso.com/addin2` ).

Comme `storage` peut être un emplacement partagé, il est important de savoir qu’il est possible de remplacer des paires clé-valeur.

Les méthodes suivantes sont disponibles avec l’objet `storage` :

 - `getItem`
 - `getItems`
 - `setItem`
 - `setItems`
 - `removeItem`
 - `removeItems`
 - `getKeys`

> [!NOTE]
> Il n’existe pas de méthode pour effacer toutes les informations (par exemple, `clear` ). À la place, vous devez utiliser l’objet `removeItems` pour supprimer plusieurs entrées à la fois.

### <a name="officeruntimestorage-example"></a>Exemple de OfficeRuntime. Storage

L’exemple de code suivant appelle la `OfficeRuntime.storage.setItem` fonction pour définir une clé et une valeur `storage` .

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

Si votre complément utilise uniquement des fonctions personnalisées sans interface utilisateur, Notez que vous ne pouvez pas accéder au modèle DOM (Document Object Model) avec des fonctions personnalisées sans interface utilisateur ou utiliser des bibliothèques telles que jQuery qui s’appuie sur le DOM.

## <a name="next-steps"></a>Étapes suivantes
Découvrez comment [Déboguer des fonctions personnalisées sans interface utilisateur](custom-functions-debugging.md).

## <a name="see-also"></a>Voir aussi

* [Authentification des fonctions personnalisées sans interface utilisateur](custom-functions-authentication.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
* [Didacticiel sur les fonctions personnalisées](../tutorials/excel-tutorial-create-custom-functions.md)
