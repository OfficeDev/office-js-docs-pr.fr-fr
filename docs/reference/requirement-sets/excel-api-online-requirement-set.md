---
title: Ensemble de conditions requises de l’API JavaScript pour Excel en ligne uniquement
description: Détails sur l’ensemble de conditions requises pour ExcelApiOnline
ms.date: 12/05/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ad2a3cd627552baeb449397fa917fe10e86ebbaf
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814151"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>Ensemble de conditions requises de l’API JavaScript pour Excel en ligne uniquement

L' `ExcelApiOnline` ensemble de conditions requises est un ensemble de conditions requises spéciales qui inclut des fonctionnalités qui sont disponibles uniquement pour Excel sur le Web. Les API de cet ensemble de conditions requises sont considérées comme des API de production (non soumises à des modifications structurelles ou comportementales non documentées) pour l’hôte Excel sur le Web. `ExcelApiOnline`sont considérés comme des API de « préversion » pour les autres plateformes (Windows, Mac, iOS) et ne sont peut-être pas pris en charge par aucune de ces plateformes.

Lorsque les API dans `ExcelApiOnline` l’ensemble de conditions requises sont prises en charge sur toutes les plateformes, elles seront ajoutées`ExcelApi 1.[NEXT]`à l’ensemble de conditions requises publié suivant (). Une fois que cette nouvelle exigence est publique, ces API seront supprimées de `ExcelApiOnline`. Imaginez qu’il s’agit d’un processus de promotion similaire, qui passe de l’aperçu à la version Release.

> [!IMPORTANT]
> `ExcelApiOnline`est un sur-ensemble du jeu de conditions requises le plus récent.

> [!IMPORTANT]
> `ExcelApiOnline 1.1`est la seule version des API en ligne uniquement. En effet, Excel sur le Web disposera toujours d’une seule version disponible pour les utilisateurs qui est la version la plus récente.

## <a name="recommended-usage"></a>Utilisation recommandée

Étant `ExcelApiOnline` donné que les API sont uniquement prises en charge par Excel sur le Web, votre complément doit vérifier si l’ensemble de conditions requises est pris en charge avant d’appeler ces API. Cela évite d’appeler une API en ligne uniquement sur une autre plateforme.

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

Une fois que l’API se trouve dans un ensemble de conditions requises entre plateformes, vous `isSetSupported` devez supprimer ou modifier la vérification. Cette opération active la fonctionnalité de votre complément sur d’autres plateformes. Veillez à tester la fonctionnalité sur ces plateformes lors de l’exécution de cette modification.

> [!IMPORTANT]
> Votre manifeste ne peut `ExcelApiOnline 1.1` pas spécifier comme condition d’activation. Il ne s’agit pas d’une valeur valide à utiliser dans l' [élément Set](../manifest/set.md).

## <a name="api-list"></a>Liste des API

Les API suivantes sont actuellement disponibles pour Excel sur le Web dans le cadre de `ExcelApiOnline 1.1` l’ensemble de conditions requises.

| Class | Champs | Description |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[mentions](/javascript/api/excel/excel.comment#mentions)|Obtient les entités (par exemple, les personnes) mentionnées dans les commentaires.|
||[richContent](/javascript/api/excel/excel.comment#richcontent)|Obtient le contenu de commentaire enrichi (par exemple, les mentions dans les commentaires). Cette chaîne n’est pas destinée à être affichée aux utilisateurs finaux. Votre complément doit uniquement l’utiliser pour analyser le contenu de commentaire enrichi.|
||[updateMentions (contentWithMentions : Excel. CommentRichContent)](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|Met à jour le contenu de commentaire avec une chaîne spécialement mise en forme et une liste de mentions.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|Obtient ou définit l’adresse de messagerie de l’entité mentionnée dans Comment.|
||[id](/javascript/api/excel/excel.commentmention#id)|Obtient ou définit l’ID de l’entité. Cela correspond à l’un des ID `CommentRichContent.richContent`dans.|
||[name](/javascript/api/excel/excel.commentmention#name)|Obtient ou définit le nom de l’entité mentionnée dans Comment.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[mentions](/javascript/api/excel/excel.commentreply#mentions)|Obtient les entités (par exemple, les personnes) mentionnées dans les commentaires.|
||[richContent](/javascript/api/excel/excel.commentreply#richcontent)|Obtient le contenu de commentaire enrichi (par exemple, les mentions dans les commentaires). Cette chaîne n’est pas destinée à être affichée aux utilisateurs finaux. Votre complément doit uniquement l’utiliser pour analyser le contenu de commentaire enrichi.|
||[updateMentions (contentWithMentions : Excel. CommentRichContent)](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|Met à jour le contenu de commentaire avec une chaîne spécialement mise en forme et une liste de mentions.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#mentions)|Tableau contenant toutes les entités (par exemple, les personnes) mentionnées dans le commentaire.|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richcontent)||
|[Range](/javascript/api/excel/excel.range)|[moveTo (destinationRange : chaîne \| de plage)](/javascript/api/excel/excel.range#moveto-destinationrange-)|Déplace les valeurs de cellule, la mise en forme et les formules de la plage actuelle à la plage de destination, en remplaçant les anciennes informations de ces cellules.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent (montant : nombre)](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|Ajuste la mise en retrait de la plage de mise en forme. La valeur de retrait est comprise entre 0 et 250 et est mesurée en caractères.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-online)
- [Version d’évaluation API JavaScript Excel](./excel-preview-apis.md)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](./excel-api-requirement-sets.md)