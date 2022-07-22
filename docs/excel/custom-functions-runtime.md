---
ms.date: 06/15/2022
description: Comprendre les fonctions personnalisées Excel qui n’utilisent pas de runtime partagé et leur runtime JavaScript spécifique.
title: Runtime JavaScript uniquement pour les fonctions personnalisées
ms.localizationpriority: medium
ms.openlocfilehash: 0d3298e95ab39f976c3fbfd5c0cc4ecdd1369721
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958410"
---
# <a name="javascript-only-runtime-for-custom-functions"></a>Runtime JavaScript uniquement pour les fonctions personnalisées

Les fonctions personnalisées qui n’utilisent pas de runtime partagé utilisent un runtime JavaScript uniquement conçu pour optimiser les performances des calculs.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

Ce runtime JavaScript permet d’accéder aux API de l’espace `OfficeRuntime` de noms qui peuvent être utilisées par les fonctions personnalisées et le volet Office (qui s’exécute dans un autre runtime) pour stocker les données.

## <a name="request-external-data"></a>Demander des données externes

Dans une fonction personnalisée, vous pouvez demander des données externes à l’aide d’une API comme [Récupérer](https://developer.mozilla.org/docs/Web/API/Fetch_API) ou de [XmlHttpRequest (XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest), une API web standard qui émet des demandes HTTP pour interagir avec les serveurs.

N’oubliez pas que les fonctions personnalisées doivent utiliser des mesures de sécurité supplémentaires lors de la création de XmlHttpRequests, nécessitant la [même stratégie d’origine](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy) et [un CORS](https://www.w3.org/TR/cors/) simple.

Une implémentation CORS simple ne peut pas utiliser de cookies et prend uniquement en charge des méthodes simples (GET, HEAD, POST). Le simple CORS accepte des en-têtes simples avec des noms de champs `Accept`, `Accept-Language`, `Content-Language`. Vous pouvez également utiliser un `Content-Type` en-tête dans CORS simple, à condition que le type de contenu soit `application/x-www-form-urlencoded`, `text/plain`ou `multipart/form-data`.

## <a name="store-and-access-data"></a>Stocker et accéder aux données

Dans une fonction personnalisée qui n’utilise pas de runtime partagé, vous pouvez stocker et accéder aux données à l’aide de l’objet [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) . L’objet `Storage` est un système de stockage persistant, non chiffré et clé-valeur qui fournit une alternative à [localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage), qui ne peut pas être utilisé par les fonctions personnalisées qui utilisent le runtime JavaScript uniquement. L’objet `Storage` offre 10 Mo de données par domaine. Les domaines peuvent être partagés par plusieurs compléments.

L’objet `Storage` est une solution de stockage partagé, ce qui signifie que plusieurs parties d’un complément peuvent accéder aux mêmes données. Par exemple, les jetons pour l’authentification utilisateur peuvent être stockés dans l’objet, car il est accessible à la `Storage` fois par une fonction personnalisée (à l’aide du runtime JavaScript uniquement) et un volet Office (à l’aide d’un runtime webview complet). De même, si deux compléments partagent le même domaine (par exemple, `www.contoso.com/addin1`, `www.contoso.com/addin2`), ils sont également autorisés à partager des informations par le biais de l’objet `Storage` . Notez que les compléments qui ont des sous-domaines différents auront des instances différentes de `Storage` (par exemple, `subdomain.contoso.com/addin1`, `differentsubdomain.contoso.com/addin2`).

Étant donné que l’objet `Storage` peut être un emplacement partagé, il est important de se rendre compte qu’il est possible de remplacer les paires clé-valeur.

Les méthodes suivantes sont disponibles sur l’objet `Storage` .

- `getItem`
- `getItems`
- `setItem`
- `setItems`
- `removeItem`
- `removeItems`
- `getKeys`

> [!NOTE]
> Il n’existe aucune méthode pour effacer toutes les informations (telles que `clear`). À la place, vous devez utiliser l’objet `removeItems` pour supprimer plusieurs entrées à la fois.

### <a name="officeruntimestorage-example"></a>Exemple OfficeRuntime.storage

L’exemple de code suivant appelle la `OfficeRuntime.storage.setItem` méthode pour définir une clé et une valeur dans `storage`.

```js
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## <a name="next-steps"></a>Étapes suivantes

Découvrez comment [déboguer des fonctions personnalisées](custom-functions-debugging.md).

## <a name="see-also"></a>Voir aussi

- [Authentification pour les fonctions personnalisées sans runtime partagé](custom-functions-authentication.md)
- [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
- [Didacticiel sur les fonctions personnalisées](../tutorials/excel-tutorial-create-custom-functions.md)
