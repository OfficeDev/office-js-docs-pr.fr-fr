---
ms.date: 01/08/2019
description: Découvrez les mises à jour les plus récentes aux fonctions Excel personnalisées.
title: Fonctions personnalisées changelog (aperçu)
localization_priority: Normal
ms.openlocfilehash: 03e4dd922ac3895e11a508f97e7ac3fa3e7b1cb0
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449274"
---
# <a name="custom-functions-changelog-preview"></a>Fonctions personnalisées changelog (aperçu)

Les fonctions personnalisées Excel est toujours en version préliminaire et qui indique que des modifications fréquentes sur le produit, y compris les modifications et la publication de nouvelles fonctionnalités. Cette changelog fournit des informations sur les modifications du produit les plus récentes.

- **7 novembre 2017 :** mise à disposition des exemples et de l’aperçu des fonctions personnalisées
- **20 novembre 2017 :** correction du bogue de compatibilité pour les utilisateurs de la version 8801 et ultérieure
- **28 novembre 2017 :** prise en charge de l’annulation sur des fonctions asynchrones (nécessite la modification des fonctions de flux)
- **7 mai 2018**: prise en charge pour Mac, Excel Online et fonctions synchrones dans les processus en cours d’exécution
- **20 septembre 2018**: prise en charge de fonctions personnalisées lors de l’exécution de JavaScript. Pour plus d’informations, voir [Runtime pour les fonctions personnalisées Excel](custom-functions-runtime.md).
- **20 octobre 2018**: avec le programme[October Insiders build](https://support.office.com/en-us/article/what-s-new-for-office-insiders-c152d1e2-96ff-4ce9-8c14-e74e13847a24), les fonctions personnalisées nécessitent désormais le paramètre « id » dans vos[métadonnées fonctions personnalisées](custom-functions-json.md) pour les versions Windows Bureau et Online. Sur Mac, ce paramètre doit être ignoré. Les fonctions personnalisées prennent également en charge maintenant des paramètres facultatifs et le `any` type return.
- **12 décembre 2018**: les fonctions personnalisées incluent désormais un moyen de découvrir l’adresse d’une cellule. Pour plus d’informations, voir[Déterminer quelle cellule a appelé votre fonction personnalisée](custom-functions-overview.md#determine-which-cell-invoked-your-custom-function).
- **Le 8 janvier 2019**: liaison méthode `CustomFunctionMapping()` a été modifié pour `CustomFunctions.associate()`. Pour plus d’informations, consultez les[Meilleures pratiques en matière de questions de fonctions personnalisées (aperçu)](custom-functions-best-practices.md).

\* pour la chaîne [Office Insider](https://products.office.com/office-insider) (anciennement appelée « Insider Fast »)

Pour obtenir la liste des problèmes connus avec le produit, voir [Problèmes Connus](custom-functions-overview.md#known-issues). 

## <a name="see-also"></a>Voir aussi

* [Vue d’ensemble des fonctions personnalisées](custom-functions-overview.md)
* [Métadonnées fonctions personnalisées](custom-functions-json.md)
* [Exécution de fonctions personnalisées Excel](custom-functions-runtime.md)
* [Meilleures pratiques de fonctions personnalisées](custom-functions-best-practices.md)
* [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
* [Débogage des fonctions personnalisées](custom-functions-debugging.md)