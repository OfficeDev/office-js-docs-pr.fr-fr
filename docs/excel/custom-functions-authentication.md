---
ms.date: 1/29/2019
description: Authentifier les utilisateurs à l’aide des fonctions personnalisées dans Excel.
title: Authentification pour les fonctions personnalisées
ms.openlocfilehash: 0e42dbc93cb545660a8dbaae5bdb48724f3b7376
ms.sourcegitcommit: 33dcf099c6b3d249811580d67ee9b790c0fdccfb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/05/2019
ms.locfileid: "29745408"
---
# <a name="authentication"></a>Authentification

Dans certains scénarios, à que votre fonction personnalisée a besoin d’authentifier l’utilisateur pour accéder des ressources protégées. Alors que les fonctions personnalisées ne nécessite pas une méthode d’authentification spécifique, sachez que les fonctions personnalisées s’exécute dans une exécution distincte dans le volet de tâches et d’autres éléments de l’interface utilisateur de votre complément. Pour cette raison, vous devez passer des données aller-retour entre les deux runtimes à l’aide de la `AsyncStorage` objet et l’API de boîte de dialogue.
  
## <a name="asyncstorage-object"></a>Objet AsyncStorage

Le module d’exécution des fonctions personnalisées ne possède pas une `localStorage` objet disponible dans la fenêtre globale, où vous pouvez généralement stocker des données. Au lieu de cela, vous devez partager des données entre des fonctions personnalisées et des volets de tâches, à l’aide de [OfficeRuntime.AsyncStorage](https://docs.microsoft.com/javascript/api/office-runtime/officeruntime.asyncstorage) pour définir et obtenir des données. 

En outre, il est un avantage à l’aide de `AsyncStorage`; Il utilise un environnement de bac à sable sécurisé afin que vos données ne sont pas accessibles par les autres compléments.  

### <a name="suggested-usage"></a>Utilisation suggérée

Lorsque vous avez besoin authentifier à partir du volet Office ou une fonction personnalisée, vérifiez AsyncStorage pour voir si le jeton d’accès a été déjà acquises. Si ce n’est pas le cas, utilisez la boîte de dialogue API pour authentifier l’utilisateur, extraire le jeton d’accès, puis enregistrez le jeton dans AsyncStorage pour une utilisation future.

## <a name="dialog-api"></a>API de boîte de dialogue

Si un jeton n’existe pas, vous devez utiliser l’API de boîte de dialogue pour demander à l’utilisateur de se connecter. Lorsqu’un utilisateur entre ses informations d’identification, le jeton d’accès qui en résulte peut être stocké dans `AsyncStorage`.

> [!NOTE]
> Le runtime de fonctions personnalisées utilise un objet Dialog est légèrement différent de l’objet Dialog dans le module d’exécution utilisé par les volets de tâches. Ils sont tous deux appelés « L’API de boîte de dialogue », mais utilisez `Officeruntime.Dialog` pour authentifier les utilisateurs dans le module d’exécution des fonctions personnalisées.

Pour plus d’informations sur l’utilisation de la `OfficeRuntime.Dialog`, voir [exécution des fonctions personnalisées](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-runtime?view=office-js#displaying-a-dialog-box).

Lors de la prévision de l’ensemble du processus d’authentification, il peut être utile de considérer le volet de tâches et les éléments d’interface utilisateur de votre complément et personnalisé fonctionne portions de votre complément en tant qu’entités distinctes qui peuvent communiquer les uns avec les autres via `AsyncStorage`.

Le diagramme suivant présente le processus de base. Notez que la ligne en pointillés indique que tout en exécutant des actions distinctes, des fonctions personnalisées et du volet de tâches de votre complément sont les deux composants de votre complément dans sa globalité.

1. Vous émettez un appel de fonction personnalisée à partir d’une cellule dans un classeur Excel.
2. Utilise la fonction personnalisée `Officeruntime.Dialog` à transmettre vos informations d’identification de l’utilisateur à un site Web.
3. Ce site Web puis renvoie un jeton d’accès à la fonction personnalisée.
4. Votre fonction personnalisée définit ensuite ce jeton d’accès la `AsyncStorage`.
5. Volet de tâches de votre complément accède au jeton de `AsyncStorage`.

![Diagramme des fonctions personnalisées, OfficeRuntime et volets collaborer.] (../images/Authdiagram.png "Diagramme d’authentification.")

## <a name="general-guidance"></a>Conseils d’ordre général

Compléments Office sont basés sur le web et vous pouvez utiliser les techniques d’authentification web. Il n’existe aucun modèle particulier ou une méthode à suivre pour implémenter votre propre authentification avec des fonctions personnalisées. Vous pouvez souhaiter, consultez la documentation sur les différents modèles d’authentification, en commençant par [cet article sur l’autorisation par le biais des services externes](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins?view=office-js).  

Évitez d’utiliser des emplacements suivants pour stocker les données lors du développement de fonctions personnalisées :  

- `localStorage`: Fonctions personnalisées n’ont pas accès au modèle global `window` objet et par conséquent, n’ont pas accès aux données stockées dans `localStorage`.
- `Office.context.document.settings`: Cet emplacement n’est pas sécurisé et informations pouvant être extraites par une personne à l’aide de la macro complémentaire.

## <a name="see-also"></a>Voir aussi

* [Métadonnées fonctions personnalisées](custom-functions-json.md)
* [Exécution de fonctions personnalisées Excel](custom-functions-runtime.md)
* [Meilleures pratiques de fonctions personnalisées](custom-functions-best-practices.md)
* [Didacticiel de fonctions personnalisées Excel](excel-tutorial-custom-functions.md)
