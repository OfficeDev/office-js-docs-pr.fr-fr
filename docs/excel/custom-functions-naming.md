---
ms.date: 12/28/2019
description: Découvrez les conditions requises pour les noms des fonctions personnalisées Excel et éviter les pièges de dénomination courants.
title: Instructions d’affectation de noms pour les fonctions personnalisées dans Excel
localization_priority: Normal
ms.openlocfilehash: 2dcd35a91f460fcd556dec479fb717942a987908
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324658"
---
# <a name="naming-guidelines"></a>Instructions d’attribution de noms

Une fonction personnalisée est identifiée par `id` une `name` propriété and dans le fichier de métadonnées JSON.

- La fonction `id` est utilisée pour identifier des fonctions personnalisées de manière unique dans votre code JavaScript.
- La fonction `name` est utilisée comme nom complet qui s’affiche pour un utilisateur dans Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Une fonction `name` peut différer de la `id`fonction, par exemple à des fins de localisation. En règle générale, les fonctions `name` d’une fonction doivent rester les `id` mêmes que s’il n’y a aucune raison impérieuse de les différencier.

Une fonction `name` et `id` partagent des exigences communes :

- Une fonction `id` ne peut utiliser que les caractères A à Z, les chiffres 0 à 9, les traits de soulignement et les points.

- Une fonction `name` peut utiliser n’importe quel caractère alphabétique Unicode, des traits de soulignement et des points.

- Les deux `name` fonctions `id` et doivent commencer par une lettre et comporter une limite minimale de trois caractères.

Excel utilise des lettres majuscules pour les noms de fonctions intégrées ( `SUM`par exemple,). Par conséquent, envisagez d’utiliser des lettres majuscules `id` pour votre fonction personnalisée et constitue `name` une meilleure pratique.

Une fonction `name` ne doit pas être nommée de la manière suivante :

- Toutes les cellules comprises entre a1 et XFD1048576 ou toutes les cellules comprises entre R1C1 et R1048576C16384.

- N’importe quelle fonction macro Excel 4,0 ( `RUN`telle `ECHO`que,).  Pour obtenir la liste complète de ces fonctions, consultez [le document de référence des fonctions de macro Excel](https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf).

## <a name="naming-conflicts"></a>Conflits de noms

Si votre fonction `name` est identique à une fonction `name` dans un complément qui existe déjà, le **#REF !** une erreur apparaît dans votre classeur.

Pour résoudre un conflit d’affectation de noms `name` , modifiez le dans votre complément et renouvelez la fonction. Vous pouvez également désinstaller le complément avec le nom conflictuel. Ou, si vous testez votre complément dans différents environnements, essayez d’utiliser un espace de noms différent pour différencier votre fonction ( `NAMESPACE_NAMEOFFUNCTION`telle que).

## <a name="best-practices"></a>Meilleures pratiques

- Envisagez d’ajouter plusieurs arguments à une fonction plutôt que de créer plusieurs fonctions avec des noms identiques ou similaires.
- Les noms de fonction doivent indiquer l’action de la fonction, `=GETZIPCODE` par exemple `ZIPCODE`au lieu de.
- Évitez les abréviations ambiguës dans les noms de fonction. La clarté est plus importante que la concision. Choisissez un nom tel `=INCREASETIME` que plutôt `=INC`que.
- Utilisez régulièrement les mêmes verbes pour les fonctions qui effectuent des actions similaires. Par exemple, utilisez `=DELETEZIPCODE` and `=DELETEADDRESS`, et non `=DELETEZIPCODE` et `=REMOVEADDRESS`.
- Lorsque vous nommez une fonction de diffusion en continu, envisagez d’ajouter une note à cet effet `STREAM` dans la description de la fonction ou d’ajouter à la fin du nom de la fonction.

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="localizing-function-names"></a>Localisation des noms de fonction

Vous pouvez localiser vos noms de fonction pour différentes langues à l’aide de fichiers JSON distincts et remplacer les valeurs dans le fichier manifeste de votre complément. Nous vous recommandons de ne pas donner à vos fonctions `id` une `name` ou une fonction Excel intégrée dans un autre langage, car cela peut entraîner des conflits avec des fonctions localisées.

Pour obtenir des informations complètes sur la localisation, voir [Localize Custom Functions](custom-functions-localize.md)

## <a name="next-steps"></a>Étapes suivantes
Découvrez les [meilleures pratiques en matière de gestion des erreurs](custom-functions-errors.md).

## <a name="see-also"></a>Voir aussi

* [Métadonnées fonctions personnalisées](custom-functions-json.md)
* [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
* [Exécution de fonctions personnalisées Excel](custom-functions-runtime.md)
