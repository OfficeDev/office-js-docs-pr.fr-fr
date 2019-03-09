---
ms.date: 02/08/2019
description: Découvrez les conditions requises pour les noms des fonctions personnalisées Excel et éviter les pièges de dénomination courants.
title: Instructions d'affectation de noms pour les fonctions personnalisées dans Excel (aperçu)
localization_priority: Normal
ms.openlocfilehash: 954753c35d2df59093661e3b8e92adfa1302e595
ms.sourcegitcommit: 8e7b7b0cfb68b91a3a95585d094cf5f5ffd00178
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/09/2019
ms.locfileid: "30512838"
---
# <a name="naming-guidelines"></a>Instructions d'affectation de noms

Une fonction personnalisée est identifiée par un **ID** et une propriété de **nom** dans le fichier de métadonnées JSON. L'ID de la fonction permet d'identifier de manière unique les fonctions personnalisées dans votre code JavaScript. Le nom de la fonction est utilisé comme nom complet qui apparaît pour un utilisateur dans Excel. Un nom de fonction peut différer de l'ID de fonction, par exemple à des fins de localisation. Toutefois, en général, il doit rester identique à l'ID s'il n'y a aucune raison impérieuse qu'ils diffèrent.

Les noms de fonction et les ID de fonction partagent des exigences communes:

- Les ID de fonction ne peuvent utiliser que les caractères A à Z, les chiffres 0 à 9, les traits de soulignement et les points.

- Les noms de fonction peuvent utiliser n'importe quel caractère alphabétique Unicode, des traits de soulignement et des points.

- Ils doivent commencer par une lettre et avoir une limite minimale de trois caractères.

Excel utilise des lettres majuscules pour les noms de fonctions intégrées ( `SUM`par exemple,). Par conséquent, il est recommandé d'utiliser des lettres majuscules pour vos noms de fonction et ID de fonction personnalisés en tant que meilleure pratique.

Les noms de fonction ne doivent pas porter le même nom que:

- Toutes les cellules comprises entre a1 et XFD1048576 ou toutes les cellules comprises entre R1C1 et R1048576C16384.

- N'importe quelle fonction macro Excel 4,0 ( `RUN`telle `ECHO`que,).  Pour obtenir une liste complète de ces fonctions, consultez [cet article](https://www.microsoft.com/en-us/download/details.aspx?id=1465).

## <a name="naming-conflicts"></a>Conflits de noms

Si le nom de votre fonction est identique à celui d'un nom de fonction dans un complément qui existe déjà, le **#REF!** une erreur apparaît dans votre classeur.

Pour résoudre un conflit de nom, modifiez le nom dans votre complément et réessayez. Vous pouvez également désinstaller le complément avec le nom conflictuel. Ou, si vous testez votre complément dans différents environnements, essayez d'utiliser un espace de noms différent pour différencier votre fonction (par exemple, NAMESPACE_NAMEOFFUNCTION).

Réfléchissez également à la façon dont vous souhaitez que les personnes utilisent les fonctions dans votre complément. Dans de nombreux cas, il est logique d'ajouter plusieurs arguments à une fonction plutôt que de créer plusieurs fonctions avec des noms identiques ou similaires.

## <a name="see-also"></a>Voir aussi

* [Métadonnées fonctions personnalisées](custom-functions-json.md)
* [Meilleures pratiques de fonctions personnalisées](custom-functions-best-practices.md)
* [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
* [Exécution de fonctions personnalisées Excel](custom-functions-runtime.md)
