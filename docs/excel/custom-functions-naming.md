---
title: Recommandations en matière d’attribution de noms pour les fonctions personnalisées dans Excel
description: Découvrez les conditions requises pour les noms Excel fonctions personnalisées et évitez les obstacles courants à l’attribution de noms.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 09216a2a98f6c37b17a413042804946bff4e6caa
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150336"
---
# <a name="custom-functions-naming-guidelines"></a>Règles de noms des fonctions personnalisées

Une fonction personnalisée est identifiée par une `id` propriété et une propriété dans le fichier de `name` métadonnées JSON.

- La fonction `id` est utilisée pour identifier de manière unique les fonctions personnalisées dans votre code JavaScript.
- La fonction `name` est utilisée comme nom d’affichage qui apparaît à l’utilisateur dans Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Une fonction peut différer de la fonction, par exemple à `name` des fins de `id` localisation. En règle générale, la fonction d’une fonction doit rester la même que s’il n’y a aucune raison qu’elle `name` `id` diffère.

Une fonction et `name` partagent `id` certaines conditions requises courantes.

- Une fonction ne peut utiliser que les caractères A à Z, les nombres de zéro à neuf, les traits de soulignement `id` et les points.

- Une fonction peut utiliser `name` n’importe quel caractère alphabétique Unicode, traits de soulignement et point.

- Les deux `name` fonctions doivent commencer par une lettre et avoir une limite minimale de trois `id` caractères.

Excel lettres majuscules pour les noms de fonctions intégrées (par `SUM` exemple). Utilisez des lettres majuscules pour votre fonction personnalisée et `name` `id` comme meilleure pratique.

Une fonction ne `name` doit pas être la même que :

- Toutes les cellules entre A1 et XFD1048576 ou les cellules entre R1C1 et R1048576C16384.

- Toute Excel fonction de macro 4.0 (telle que `RUN` , `ECHO` ).  Pour obtenir la liste complète de ces fonctions, consultez ce document Excel référence des fonctions de [macro.](https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf)

## <a name="naming-conflicts"></a>Conflits d’attribution de noms

Si votre fonction est identique à une fonction dans un module qui existe déjà dans un `name` `name` **#REF!** s’affiche dans votre workbook.

Pour résoudre un conflit d’attribution de noms, modifiez le nom de votre `name` add-in et réessayez la fonction. Vous pouvez également désinstaller le add-in avec le nom en conflit. Ou, si vous testez votre add-in dans différents environnements, essayez d’utiliser un espace de noms différent pour différencier votre fonction (par `NAMESPACE_NAMEOFFUNCTION` exemple).

## <a name="best-practices"></a>Meilleures pratiques

- Envisagez d’ajouter plusieurs arguments à une fonction plutôt que de créer plusieurs fonctions avec des noms identiques ou similaires.
- Évitez les abréviations ambiguës dans les noms de fonctions. La clarté est plus importante que la concision. Choisissez un nom comme `=INCREASETIME` plutôt que `=INC` .
- Les noms de fonction doivent indiquer l’action de la fonction, telle que =GETZIPCODE au lieu de ZIPCODE.
- Utilisez systématiquement les mêmes verbes pour les fonctions qui effectuent des actions similaires. Par exemple, utilisez `=DELETEZIPCODE` et , plutôt que et `=DELETEADDRESS` `=DELETEZIPCODE` `=REMOVEADDRESS` .
- Lorsque vous nommez une fonction de diffusion en continu, envisagez d’ajouter une note à cet effet dans la description de la fonction ou d’ajouter à la fin du nom `STREAM` de la fonction.

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="localizing-function-names"></a>Localisation des noms de fonctions

Vous pouvez localiser les noms de vos fonctions pour différentes langues à l’aide de fichiers JSON distincts et de valeurs de remplacement dans le fichier manifeste de votre add-in. Évitez d’accorder à vos fonctions une fonction Excel dans un autre langage, car cela peut être en conflit avec `id` `name` des fonctions localisées.

Pour plus d’informations sur la recherche, voir [Localize custom functions](custom-functions-localize.md)

## <a name="next-steps"></a>Étapes suivantes

Découvrez les meilleures pratiques [de gestion des erreurs.](custom-functions-errors.md)

## <a name="see-also"></a>Voir aussi

* [Créer manuellement des métadonnées JSON pour les fonctions personnalisées](custom-functions-json.md)
* [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
