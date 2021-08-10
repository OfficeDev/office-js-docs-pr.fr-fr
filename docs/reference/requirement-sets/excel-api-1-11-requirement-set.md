---
title: Excel Ensemble de conditions requises de l’API JavaScript 1.11
description: Détails sur l’ensemble de conditions requises ExcelApi 1.11.
ms.date: 04/01/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 67fb212813608ecb4e72ba5d63952f0228875211d0bf66978b7201fff58c5076
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092655"
---
# <a name="whats-new-in-excel-javascript-api-111"></a>Nouveautés de l Excel API JavaScript 1.11

ExcelApi 1.11 a amélioré la prise en charge des commentaires et des contrôles au niveau du workbook (par exemple, l’enregistrement et la fermeture du livre). Il a également ajouté l’accès aux paramètres de culture pour prendre en compte la localisation.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| Mentions [de commentaire](../../excel/excel-add-ins-comments.md#mentions) |Balise et avertit les autres utilisateurs du classez par le biais de commentaires. | [Comment](/javascript/api/excel/excel.comment), [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) |
| Résolution des [commentaires](../../excel/excel-add-ins-comments.md#resolve-comment-threads) | Résolvez les threads de commentaires et obtenez l’état de résolution. | [Commentaire](/javascript/api/excel/excel.comment) |
| [Paramètres de culture](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | Obtient les paramètres du système culturel pour le workbook, tels que la mise en forme des nombres. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [Application NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [](/javascript/api/excel/excel.application) |
| [Couper et coller (moveTo)](../../excel/excel-add-ins-ranges-cut-copy-paste.md) | Réplique la fonctionnalité couper-coller dans Excel pour une plage. | [Range](/javascript/api/excel/excel.range) |
| Classeur [enregistrer](../../excel/excel-add-ins-workbooks.md#save-the-workbook) et [fermer](../../excel/excel-add-ins-workbooks.md#close-the-workbook) | Enregistrez et fermez ses classeurs. | [Workbook](/javascript/api/excel/excel.workbook) |
| Événements de feuille de calcul | Informations supplémentaires sur les événements et les événements pour les calculs de feuille de calcul et les lignes masquées. | [WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs), [WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API Excel l’ensemble de conditions requises de l’API JavaScript 1.11. Pour afficher la documentation de référence de l’API pour toutes les API prise en charge par Excel l’ensemble de conditions requises de l’API JavaScript 1.11 ou une version antérieure, voir les API Excel dans l’ensemble de conditions requises [1.11](/javascript/api/excel?view=excel-js-1.11&preserve-view=true)ou version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureInfo)|Fournit des informations basées sur les paramètres de culture système actuels.|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalSeparator)|Obtient la chaîne utilisée comme séparateur décimal pour les valeurs numériques.|
||[thousandsSeparator](/javascript/api/excel/excel.application#thousandsSeparator)|Obtient la chaîne utilisée pour séparer les groupes de chiffres à gauche de la virgule pour les valeurs numériques.|
||[useSystemSeparators](/javascript/api/excel/excel.application#useSystemSeparators)|Spécifie si les séparateurs système de Excel sont activés.|
|[Commentaire](/javascript/api/excel/excel.comment)|[mentions](/javascript/api/excel/excel.comment#mentions)|Obtient les entités (par exemple, les personnes) mentionnées dans les commentaires.|
||[richContent](/javascript/api/excel/excel.comment#richContent)|Obtient le contenu des commentaires enrichis (par exemple, les mentions dans les commentaires).|
||[résolu](/javascript/api/excel/excel.comment#resolved)|État du thread de commentaire.|
||[updateMentions(contentWithMentions: Excel. CommentRichContent)](/javascript/api/excel/excel.comment#updateMentions_contentWithMentions_)|Met à jour le contenu des commentaires avec une chaîne spécialement mise en forme et une liste de mentions.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: CommentRichContent \| string, contentType?: Excel. ContentType)](/javascript/api/excel/excel.commentcollection#add_cellAddress__content__contentType_)|Crée un nouveau commentaire avec le contenu donné sur la cellule donnée.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|Adresse de messagerie de l’entité mentionnée dans un commentaire.|
||[id](/javascript/api/excel/excel.commentmention#id)|ID de l’entité.|
||[name](/javascript/api/excel/excel.commentmention#name)|Nom de l’entité mentionnée dans un commentaire.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[mentions](/javascript/api/excel/excel.commentreply#mentions)|Entités (par exemple, personnes) mentionnées dans les commentaires.|
||[résolu](/javascript/api/excel/excel.commentreply#resolved)|État de réponse du commentaire.|
||[richContent](/javascript/api/excel/excel.commentreply#richContent)|Contenu de commentaire enrichi (par exemple, mentions dans les commentaires).|
||[updateMentions(contentWithMentions: Excel. CommentRichContent)](/javascript/api/excel/excel.commentreply#updateMentions_contentWithMentions_)|Met à jour le contenu des commentaires avec une chaîne spécialement mise en forme et une liste de mentions.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: CommentRichContent \| string, contentType?: Excel. ContentType)](/javascript/api/excel/excel.commentreplycollection#add_content__contentType_)|Crée une réponse de commentaire pour un commentaire.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#mentions)|Tableau contenant toutes les entités (par exemple, les personnes) mentionnées dans le commentaire.|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richContent)|Spécifie le contenu enrichi du commentaire (par exemple, le contenu de commentaire avec mentions, la première entité mentionnée a un attribut d’ID de 0 et la seconde entité mentionnée a un attribut d’ID de 1).|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#name)|Obtient le nom de la culture au format languagecode2-country/regioncode2 (par exemple, « zh-cn » ou « en-us »).|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#numberFormat)|Définit le format adapté à la culture de l’affichage des nombres.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberDecimalSeparator)|Obtient la chaîne utilisée comme séparateur décimal pour les valeurs numériques.|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#numberGroupSeparator)|Obtient la chaîne utilisée pour séparer les groupes de chiffres à gauche de la virgule pour les valeurs numériques.|
|[Range](/javascript/api/excel/excel.range)|[moveTo(destinationRange: Range \| string)](/javascript/api/excel/excel.range#moveTo_destinationRange_)|Déplace les valeurs, la mise en forme et les formules des cellules de la plage actuelle vers la plage de destination, en remplaçant les anciennes informations de ces cellules.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent(amount: number)](/javascript/api/excel/excel.rangeformat#adjustIndent_amount_)|Ajuste le retrait de la mise en forme de plage.|
|[Classeur](/javascript/api/excel/excel.workbook)|[Fermer (closeBehavior ? : Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close_closeBehavior_)|Fermer le classeur actif.|
||[Enregistrer (saveBehavior ? : Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save_saveBehavior_)|Enregistrer le classeur actif.|
|[Feuille de calcul](/javascript/api/excel/excel.worksheet)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onRowHiddenChanged)|Se produit lorsque l’état masqué d’une ou plusieurs lignes a changé dans une feuille de calcul spécifique.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[adresse](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|Adresse de la plage qui a effectué le calcul.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onRowHiddenChanged)|Se produit lorsque l’état masqué d’une ou plusieurs lignes a changé dans une feuille de calcul spécifique.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[adresse](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changeType)|Obtient le type de modification qui représente la façon dont l’événement a été déclenché.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetId)|Obtient l’ID de la feuille de calcul dans laquelle les données ont été modifiées.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.11&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
