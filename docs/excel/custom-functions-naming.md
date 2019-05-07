---
ms.date: 05/03/2019
description: Découvrez les conditions requises pour les noms des fonctions personnalisées Excel et éviter les pièges de dénomination courants.
title: Instructions d’affectation de noms pour les fonctions personnalisées dans Excel
localization_priority: Normal
ms.openlocfilehash: 3abe04eebfa703666b70ecbde1c68ab0c942003c
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628045"
---
# <a name="naming-guidelines"></a>Instructions d’affectation de noms

Une fonction personnalisée est identifiée par un **ID** et une propriété de **nom** dans le fichier de métadonnées JSON.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

- La fonction `id` est utilisée pour identifier des fonctions personnalisées de manière unique dans votre code JavaScript. 
- La fonction `name` est utilisée comme nom complet qui s’affiche pour un utilisateur dans Excel. 

Une fonction `name` peut différer de la `id`fonction, par exemple à des fins de localisation. En règle générale, les fonctions `name` d’une fonction doivent rester les `id` mêmes que s’il n’y a aucune raison impérieuse de les différencier.

Une fonction `name` et `id` partagent des exigences communes:

- Une fonction `id` ne peut utiliser que les caractères A à Z, les chiffres 0 à 9, les traits de soulignement et les points.

- Une fonction `name` peut utiliser n’importe quel caractère alphabétique Unicode, des traits de soulignement et des points.

- Les deux `name` fonctions `id` et doivent commencer par une lettre et comporter une limite minimale de trois caractères.

Excel utilise des lettres majuscules pour les noms de fonctions intégrées ( `SUM`par exemple,). Par conséquent, envisagez d’utiliser des lettres majuscules `id` pour votre fonction personnalisée et constitue `name` une meilleure pratique.

Une fonction `name` ne doit pas être nommée de la manière suivante:

- Toutes les cellules comprises entre a1 et XFD1048576 ou toutes les cellules comprises entre R1C1 et R1048576C16384.

- N’importe quelle fonction macro Excel 4,0 ( `RUN`telle `ECHO`que,).  Pour obtenir une liste complète de ces fonctions, consultez [cet article](https://www.microsoft.com/en-us/download/details.aspx?id=1465).

## <a name="naming-conflicts"></a>Conflits de noms

Si votre fonction `name` est identique à une fonction `name` dans un complément qui existe déjà, le **#REF!** une erreur apparaît dans votre classeur.

Pour résoudre un conflit d’affectation de noms `name` , modifiez le dans votre complément et renouvelez la fonction. Vous pouvez également désinstaller le complément avec le nom conflictuel. Ou, si vous testez votre complément dans différents environnements, essayez d’utiliser un espace de noms différent pour différencier votre fonction ( `NAMESPACE_NAMEOFFUNCTION`telle que).

## <a name="best-practices"></a>Meilleures pratiques

- Envisagez d’ajouter plusieurs arguments à une fonction plutôt que de créer plusieurs fonctions avec des noms identiques ou similaires.
- Les noms de fonction doivent indiquer l’action de la fonction, `=GETZIPCODE` par exemple `ZIPCODE`au lieu de.
- Évitez les abréviations ambiguës dans les noms de fonction. La clarté est plus importante que la concision. Choisissez un nom tel `=INCREASETIME` que plutôt `=INC`que.
- Utilisez régulièrement les mêmes verbes pour les fonctions qui effectuent des actions similaires. Par exemple, utilisez `=DELETEZIPCODE` and `=DELETEADDRESS`, et non `=DELETEZIPCODE` et `=REMOVEADDRESS`.

## <a name="localizing-function-names"></a>Localisation des noms de fonction

Vous pouvez localiser vos noms de fonction pour différentes langues à l’aide de fichiers JSON distincts et remplacer les valeurs dans le fichier manifeste de votre complément. Nous vous recommandons de ne pas donner à vos fonctions `id` une `name` ou une fonction Excel intégrée dans un autre langage, car cela peut entraîner des conflits avec des fonctions localisées.

Pour obtenir des informations complètes sur la localisation, voir [Localize Custom Functions](custom-functions-localize.md)

## <a name="next-steps"></a>Étapes suivantes
Découvrez les [meilleures pratiques en matière de gestion des erreurs](custom-functions-errors.md).

## <a name="see-also"></a>Voir aussi

* [Métadonnées fonctions personnalisées](custom-functions-json.md)
* [Meilleures pratiques de fonctions personnalisées](custom-functions-best-practices.md)
* [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
* [Exécution de fonctions personnalisées Excel](custom-functions-runtime.md)
