---
title: Excel conditions requises de l’API JavaScript 1.11
description: Détails sur l’ensemble de conditions requises ExcelApi 1.11.
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
---

# <a name="whats-new-in-excel-javascript-api-111"></a>Nouveautés de l Excel API JavaScript 1.11

ExcelApi 1.11 a amélioré la prise en charge des commentaires et des contrôles au niveau du workbook (par exemple, l’enregistrement et la fermeture du manuel). Il a également ajouté l’accès aux paramètres de culture pour prendre en compte la localisation.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| [Mentions de commentaire](../../excel/excel-add-ins-comments.md#mentions) |Balise et avertit les autres utilisateurs du classez par le biais de commentaires. | [Comment](/javascript/api/excel/excel.comment), [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) |
| Résolution des [commentaires](../../excel/excel-add-ins-comments.md#resolve-comment-threads) | Résolvez les threads de commentaires et obtenez l’état de résolution. | [Comment](/javascript/api/excel/excel.comment) |
| [Paramètres de culture](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | Obtient les paramètres du système culturel pour le workbook, tels que la mise en forme des nombres. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [Application NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [](/javascript/api/excel/excel.application) |
| [Couper et coller (moveTo)](../../excel/excel-add-ins-ranges-cut-copy-paste.md) | Réplique la fonctionnalité couper-coller dans Excel pour une plage. | [Range](/javascript/api/excel/excel.range) |
| Classeur [enregistrer](../../excel/excel-add-ins-workbooks.md#save-the-workbook) et [fermer](../../excel/excel-add-ins-workbooks.md#close-the-workbook) | Enregistrez et fermez ses classeurs. | [Workbook](/javascript/api/excel/excel.workbook) |
| Événements de feuille de calcul | Informations supplémentaires sur les événements et les événements pour les calculs de feuille de calcul et les lignes masquées. | [WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs), [WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API de Excel l’ensemble de conditions requises de l’API JavaScript 1.11. Pour afficher la documentation de référence de l’API pour toutes les API prise en charge par Excel l’ensemble de conditions requises de l’API JavaScript 1.11 ou une version antérieure, voir les API Excel dans l’ensemble de conditions requises [1.11](/javascript/api/excel?view=excel-js-1.11&preserve-view=true) ou une version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#excel-excel-application-cultureinfo-member)|Fournit des informations basées sur les paramètres de culture système actuels.|
||[decimalSeparator](/javascript/api/excel/excel.application#excel-excel-application-decimalseparator-member)|Obtient la chaîne utilisée comme séparateur décimal pour les valeurs numériques.|
||[thousandsSeparator](/javascript/api/excel/excel.application#excel-excel-application-thousandsseparator-member)|Obtient la chaîne utilisée pour séparer les groupes de chiffres à gauche de la virgule pour les valeurs numériques.|
||[useSystemSeparators](/javascript/api/excel/excel.application#excel-excel-application-usesystemseparators-member)|Spécifie si les séparateurs système de Excel sont activés.|
|[Comment](/javascript/api/excel/excel.comment)|[mentions](/javascript/api/excel/excel.comment#excel-excel-comment-mentions-member)|Obtient les entités (par exemple, les personnes) mentionnées dans les commentaires.|
||[résolu](/javascript/api/excel/excel.comment#excel-excel-comment-resolved-member)|État du thread de commentaire.|
||[richContent](/javascript/api/excel/excel.comment#excel-excel-comment-richcontent-member)|Obtient le contenu des commentaires enrichis (par exemple, les mentions dans les commentaires).|
||[updateMentions(contentWithMentions: Excel. CommentRichContent)](/javascript/api/excel/excel.comment#excel-excel-comment-updatementions-member(1))|Met à jour le contenu des commentaires avec une chaîne spécialement mise en forme et une liste de mentions.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: CommentRichContent \| string, contentType?: Excel. ContentType)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-add-member(1))|Crée un nouveau commentaire avec le contenu donné sur la cellule donnée.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#excel-excel-commentmention-email-member)|Adresse de messagerie de l’entité mentionnée dans un commentaire.|
||[id](/javascript/api/excel/excel.commentmention#excel-excel-commentmention-id-member)|ID de l’entité.|
||[name](/javascript/api/excel/excel.commentmention#excel-excel-commentmention-name-member)|Nom de l’entité mentionnée dans un commentaire.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[mentions](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-mentions-member)|Entités (par exemple, personnes) mentionnées dans les commentaires.|
||[résolu](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-resolved-member)|État de réponse du commentaire.|
||[richContent](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-richcontent-member)|Contenu de commentaire enrichi (par exemple, mentions dans les commentaires).|
||[updateMentions(contentWithMentions: Excel. CommentRichContent)](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-updatementions-member(1))|Met à jour le contenu des commentaires avec une chaîne spécialement mise en forme et une liste de mentions.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: CommentRichContent \| string, contentType?: Excel. ContentType)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-add-member(1))|Crée une réponse de commentaire pour un commentaire.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#excel-excel-commentrichcontent-mentions-member)|Tableau contenant toutes les entités (par exemple, les personnes) mentionnées dans le commentaire.|
||[richContent](/javascript/api/excel/excel.commentrichcontent#excel-excel-commentrichcontent-richcontent-member)|Spécifie le contenu enrichi du commentaire (par exemple, le contenu de commentaire avec mentions, la première entité mentionnée a un attribut d’ID de 0 et la seconde entité mentionnée a un attribut d’ID de 1).|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#excel-excel-cultureinfo-name-member)|Obtient le nom de la culture au format languagecode2-country/regioncode2 (par exemple, « zh-cn » ou « en-us »).|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#excel-excel-cultureinfo-numberformat-member)|Définit le format adapté à la culture de l’affichage des nombres.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#excel-excel-numberformatinfo-numberdecimalseparator-member)|Obtient la chaîne utilisée comme séparateur décimal pour les valeurs numériques.|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#excel-excel-numberformatinfo-numbergroupseparator-member)|Obtient la chaîne utilisée pour séparer les groupes de chiffres à gauche de la virgule pour les valeurs numériques.|
|[Range](/javascript/api/excel/excel.range)|[moveTo(destinationRange: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-moveto-member(1))|Déplace les valeurs, la mise en forme et les formules des cellules de la plage actuelle vers la plage de destination, en remplaçant les anciennes informations de ces cellules.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent(amount: number)](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-adjustindent-member(1))|Ajuste le retrait de la mise en forme de plage.|
|[Workbook](/javascript/api/excel/excel.workbook)|[Fermer (closeBehavior ? : Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#excel-excel-workbook-close-member(1))|Fermer le classeur actif.|
||[Enregistrer (saveBehavior ? : Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#excel-excel-workbook-save-member(1))|Enregistrer le classeur actif.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onrowhiddenchanged-member)|Se produit lorsque l’état masqué d’une ou plusieurs lignes a changé dans une feuille de calcul spécifique.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[adresse](/javascript/api/excel/excel.worksheetcalculatedeventargs#excel-excel-worksheetcalculatedeventargs-address-member)|Adresse de la plage qui a effectué le calcul.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onrowhiddenchanged-member)|Se produit lorsque l’état masqué d’une ou plusieurs lignes a changé dans une feuille de calcul spécifique.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[adresse](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-address-member)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-changetype-member)|Obtient le type de modification qui représente la façon dont l’événement a été déclenché.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-source-member)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-type-member)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-worksheetid-member)|Obtient l’ID de la feuille de calcul dans laquelle les données ont été modifiées.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.11&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
