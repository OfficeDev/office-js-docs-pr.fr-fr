---
ms.date: 01/29/2019
description: Authentifier les utilisateurs à l'aide de fonctions personnalisées dans Excel.
title: Authentification pour les fonctions personnalisées
ms.openlocfilehash: 260f15c39758b82a2145474f543c3c9ff5edd132
ms.sourcegitcommit: 70ef38a290c18a1d1a380fd02b263470207a5dc6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/15/2019
ms.locfileid: "30052734"
---
# <a name="authentication"></a>Authentification

Dans certains scénarios, votre fonction personnalisée doit authentifier l'utilisateur afin d'accéder aux ressources protégées. Bien que les fonctions personnalisées ne nécessitent pas une méthode d'authentification spécifique, vous devez savoir que les fonctions personnalisées s'exécutent dans un Runtime distinct à partir du volet Office et d'autres éléments d'interface utilisateur de votre complément. Pour cette raison, vous devez transmettre les données entre les deux runtimes à l'aide de l' `AsyncStorage` objet et de l'API Dialog.
  
## <a name="asyncstorage-object"></a>Objet Dansasyncstorage

Le runtime des fonctions personnalisées ne `localStorage` dispose pas d'un objet disponible dans la fenêtre globale, dans laquelle vous pouvez généralement stocker des données. Au lieu de cela, vous devez partager des données entre des fonctions personnalisées et des volets Office à l'aide de [OfficeRuntime. dansasyncstorage](https://docs.microsoft.com/javascript/api/office-runtime/officeruntime.asyncstorage) pour définir et obtenir des données. 

Par ailleurs, il est intéressant d'utiliser `AsyncStorage`; Il utilise un environnement de bac à sable (sandbox) sécurisé afin que les autres compléments ne puissent pas accéder à vos données.  

### <a name="suggested-usage"></a>Utilisation suggérée

Lorsque vous devez vous authentifier à partir du volet Office ou d'une fonction personnalisée, vérifiez Dansasyncstorage pour voir si le jeton d'accès a déjà été acquis. Si ce n'est pas le cas, utilisez l'API de boîte de dialogue pour authentifier l'utilisateur, récupérer le jeton d'accès, puis stocker le jeton dans Dansasyncstorage pour une utilisation ultérieure.

## <a name="dialog-api"></a>API de dialogue

Si un jeton n'existe pas, vous devez utiliser l'API de boîte de dialogue pour demander à l'utilisateur de se connecter. Une fois qu'un utilisateur a entré ses informations d'identification, le jeton d'accès `AsyncStorage`résultant peut être stocké dans.

> [!NOTE]
> Le runtime des fonctions personnalisées utilise un objet Dialog légèrement différent de l'objet Dialog dans le runtime utilisé par les volets des tâches. Ils sont tous deux appelés «API de dialogue», mais utilisent `Officeruntime.Dialog` pour authentifier les utilisateurs dans le runtime des fonctions personnalisées.

Pour plus d'informations sur l'utilisation `OfficeRuntime.Dialog`du, voir [Custom Functions Runtime](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-runtime?view=office-js#displaying-a-dialog-box).

Lors de l'identification de l'ensemble du processus d'authentification, il peut être utile de considérer le volet Office et les éléments de l'interface utilisateur de votre complément, ainsi que les parties fonctions personnalisées de votre complément en tant qu'entités distinctes pouvant communiquer les `AsyncStorage`uns avec les autres.

Le diagramme suivant décrit ce processus de base. Notez que la ligne pointillée indique que lorsqu'ils effectuent des actions distinctes, les fonctions personnalisées et le volet Office de votre complément sont des parties de votre complément dans son ensemble.

1. Vous émettez un appel de fonction personnalisée à partir d'une cellule dans un classeur Excel.
2. La fonction personnalisée utilise `Officeruntime.Dialog` pour transmettre les informations d'identification de votre utilisateur à un site Web.
3. Ce site Web renvoie ensuite un jeton d'accès à la fonction personnalisée.
4. Votre fonction personnalisée définit ensuite le jeton d'accès sur `AsyncStorage`le.
5. Le volet Office de votre complément accède au jeton à partir de `AsyncStorage`.

![Diagramme de fonctions personnalisées, d'OfficeRuntime et de volets de tâches qui fonctionnent ensemble.] (../images/Authdiagram.png "Diagramme d'authentification.")

## <a name="general-guidance"></a>Conseils généraux

Les compléments Office sont basés sur le Web et vous pouvez utiliser n'importe quelle technique d'authentification Web. Il n'existe pas de modèle ni de méthode particulier à respecter pour implémenter votre propre authentification avec des fonctions personnalisées. Vous pouvez consulter la documentation sur les différents modèles d'authentification, en commençant par [cet article sur la création via des services externes](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins?view=office-js).  

Évitez d'utiliser les emplacements suivants pour stocker des données lors du développement de fonctions personnalisées:  

- `localStorage`: Les fonctions personnalisées n'ont pas accès à `window` l'objet global et, par conséquent, n'ont `localStorage`pas accès aux données stockées dans.
- `Office.context.document.settings`: Cet emplacement n'est pas sécurisé et les informations peuvent être extraites par quiconque utilisant le complément.

## <a name="see-also"></a>Voir aussi

* [Métadonnées fonctions personnalisées](custom-functions-json.md)
* [Exécution de fonctions personnalisées Excel](custom-functions-runtime.md)
* [Meilleures pratiques de fonctions personnalisées](custom-functions-best-practices.md)
* [Didacticiel de fonctions personnalisées Excel](excel-tutorial-custom-functions.md)
