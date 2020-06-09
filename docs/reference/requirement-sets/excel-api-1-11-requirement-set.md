---
title: Ensemble de conditions requises de l’API JavaScript pour Excel 1,11
description: Détails sur l’ensemble de conditions requises ExcelApi 1,11
ms.date: 05/06/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ab9fde262640aa243aaf2b88767225505e08b3b7
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612093"
---
# <a name="whats-new-in-excel-javascript-api-111"></a>Nouveautés de l’API JavaScript pour Excel 1,11

La prise en charge ExcelApi 1,11 améliorée pour les commentaires et les contrôles au niveau du classeur (par exemple, l’enregistrement et la fermeture du classeur). Elle a également ajouté l’accès aux paramètres de culture pour faciliter la localisation.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| [Mentions](../../excel/excel-add-ins-comments.md#mentions) de commentaires |Balises et avertit d’autres utilisateurs du classeur par le biais de commentaires. | [Commentaire](/javascript/api/excel/excel.comment), [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) |
| [Résolution](../../excel/excel-add-ins-comments.md#resolve-comment-threads) des commentaires | Résoudre les threads de commentaires et obtenir l’état de résolution. | [Comment](/javascript/api/excel/excel.comment) |
| [Paramètres de culture](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | Obtient les paramètres du système culturel pour le classeur, tels que la mise en forme des nombres. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [application](/javascript/api/excel/excel.application) NumberFormatInfo |
| [Couper-coller (moveTo)](../../excel/excel-add-ins-ranges-advanced.md#cut-copy-and-paste) | Réplique la fonctionnalité de couper-coller dans Excel pour une plage. | [Range](/javascript/api/excel/excel.range) |
| Classeur [enregistrer](../../excel/excel-add-ins-workbooks.md#save-the-workbook) et [fermer](../../excel/excel-add-ins-workbooks.md#close-the-workbook) | Enregistrez et fermez ses classeurs. | [Workbook](/javascript/api/excel/excel.workbook) |
| Événements de feuille de calcul | Événements et informations d’événements supplémentaires pour les calculs de feuille de calcul et les lignes masquées. | [WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs), [WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API dans l’ensemble de conditions requises de l’API JavaScript pour Excel 1,11. Pour afficher la documentation de référence de l’API pour toutes les API prises en charge par l’ensemble de conditions requises de l’API JavaScript pour Excel 1,11 ou antérieure, voir [API Excel dans l’ensemble de conditions requises 1,11 ou version antérieure](/javascript/api/excel?view=excel-js-1.11).

| Class | Champs | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureinfo)|Fournit des informations basées sur les paramètres de culture système actuels. Cela inclut les noms de culture, la mise en forme de numéros et d’autres paramètres dépendants de la culture.|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalseparator)|Obtient la chaîne utilisée comme séparateur décimal pour les valeurs numériques. Cette fonction est basée sur les paramètres locaux d’Excel.|
||[thousandsSeparator](/javascript/api/excel/excel.application#thousandsseparator)|Obtient la chaîne utilisée pour séparer les groupes de chiffres à gauche du séparateur décimal pour les valeurs numériques. Cette fonction est basée sur les paramètres locaux d’Excel.|
||[UseSystemSeparators,](/javascript/api/excel/excel.application#usesystemseparators)|Indique si les séparateurs système d’Excel sont activés.|
|[Comment](/javascript/api/excel/excel.comment)|[mentions](/javascript/api/excel/excel.comment#mentions)|Obtient les entités (par exemple, les personnes) mentionnées dans les commentaires.|
||[richContent](/javascript/api/excel/excel.comment#richcontent)|Obtient le contenu de commentaire enrichi (par exemple, mentions dans les commentaires). Cette chaîne n’est pas destinée à être affichée aux utilisateurs finaux. Votre complément doit uniquement l’utiliser pour analyser le contenu de commentaire enrichi.|
||[évaluation](/javascript/api/excel/excel.comment#resolved)|État du fil de commentaire. La valeur « true » signifie que le thread de commentaire est résolu.|
||[updateMentions (contentWithMentions : Excel. CommentRichContent)](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|Met à jour le contenu de commentaire avec une chaîne spécialement mise en forme et une liste de mentions.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[Add (cellAddress : Range \| String, content : CommentRichContent \| String, ContentType ?: Excel. ContentType)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|Crée un nouveau commentaire avec le contenu donné sur la cellule donnée. Une `InvalidArgument` erreur est générée si la plage fournie est plus grande qu’une cellule.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|Adresse de messagerie de l’entité mentionnée dans Comment.|
||[id](/javascript/api/excel/excel.commentmention#id)|ID de l’entité. L’ID correspond à l’un des ID dans `CommentRichContent.richContent` .|
||[name](/javascript/api/excel/excel.commentmention#name)|Nom de l’entité mentionnée dans Comment.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[mentions](/javascript/api/excel/excel.commentreply#mentions)|Entités (par exemple, les personnes) mentionnées dans les commentaires.|
||[évaluation](/javascript/api/excel/excel.commentreply#resolved)|État de la réponse de commentaire. La valeur « true » signifie que la réponse est à l’État résolu.|
||[richContent](/javascript/api/excel/excel.commentreply#richcontent)|Contenu de commentaire enrichi (par exemple, mentions dans les commentaires). Cette chaîne n’est pas destinée à être affichée aux utilisateurs finaux. Votre complément doit uniquement l’utiliser pour analyser le contenu de commentaire enrichi.|
||[updateMentions (contentWithMentions : Excel. CommentRichContent)](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|Met à jour le contenu de commentaire avec une chaîne spécialement mise en forme et une liste de mentions.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[Add (Content : CommentRichContent \| String, ContentType ?: Excel. ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Crée une réponse à un commentaire pour un commentaire.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#mentions)|Tableau contenant toutes les entités (par exemple, les personnes) mentionnées dans le commentaire.|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richcontent)|Spécifie le contenu enrichi du commentaire (par exemple, le contenu de commentaire avec des mentions, la première entité mentionnée a un attribut ID de 0 et la deuxième entité mentionnée a un attribut ID de 1.|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#name)|Obtient le nom de la culture au format languagecode2-Country/regioncode2 (par exemple, « zh-CN » ou « en-US »). Cette fonction est basée sur les paramètres système actuels.|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#numberformat)|Définit le format d’affichage des nombres approprié pour la culture. Cette fonction est basée sur les paramètres de culture actuelle du système.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|Obtient la chaîne utilisée comme séparateur décimal pour les valeurs numériques. Cette fonction est basée sur les paramètres système actuels.|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|Obtient la chaîne utilisée pour séparer les groupes de chiffres à gauche du séparateur décimal pour les valeurs numériques. Cette fonction est basée sur les paramètres système actuels.|
|[Range](/javascript/api/excel/excel.range)|[moveTo (destinationRange : chaîne de plage \| )](/javascript/api/excel/excel.range#moveto-destinationrange-)|Déplace les valeurs de cellule, la mise en forme et les formules de la plage actuelle à la plage de destination, en remplaçant les anciennes informations de ces cellules.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent (montant : nombre)](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|Ajuste la mise en retrait de la plage de mise en forme. La valeur de retrait est comprise entre 0 et 250 et est mesurée en caractères.|
|[Workbook](/javascript/api/excel/excel.workbook)|[Fermer (closeBehavior ? : Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Fermer le classeur actif.|
||[Enregistrer (saveBehavior ? : Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Enregistrer le classeur actif.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Survient lorsque l’état masqué d’une ou plusieurs lignes a été modifié sur une feuille de calcul spécifique.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[adresse](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|Adresse de la plage qui a terminé le calcul.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Survient lorsque l’état masqué d’une ou plusieurs lignes a été modifié sur une feuille de calcul spécifique.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[adresse](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Obtient le type de modification qui représente la manière dont l’événement a été déclenché. `Excel.RowHiddenChangeType`Pour plus d’informations, voir.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle les données sont modifiées.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.11)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
