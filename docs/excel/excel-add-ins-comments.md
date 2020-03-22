---
title: Utiliser des commentaires à l’aide de l’API JavaScript pour Excel
description: Informations sur l’utilisation des API pour ajouter, supprimer et modifier des commentaires et des thèmes de commentaires.
ms.date: 03/17/2020
localization_priority: Normal
ms.openlocfilehash: 275828915730d3438101315ee28bf76aa8b8bf3f
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890569"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a>Utiliser des commentaires à l’aide de l’API JavaScript pour Excel

Cet article explique comment ajouter, lire, modifier et supprimer des commentaires dans un classeur à l’aide de l’API JavaScript pour Excel. Pour en savoir plus sur la fonctionnalité de commentaire, consultez l’article [Insérer des commentaires et des notes dans Excel](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) .

Dans l’API JavaScript pour Excel, un commentaire inclut à la fois le commentaire initial unique et la discussion liée au thread. Elle est liée à une cellule individuelle. Toute personne qui consulte le classeur avec des autorisations suffisantes peut répondre à un commentaire. Un objet [Comment](/javascript/api/excel/excel.comment) stocke ces réponses en tant qu’objets [CommentReply](/javascript/api/excel/excel.commentreply) . Vous devez considérer un commentaire comme un fil de discussion et qu’un thread doit avoir une entrée spéciale comme point de départ.

![Commentaire Excel, étiqueté « commentaire » avec deux réponses, intitulées « comment. réponses [0] » et «comment. réponses [1].](../images/excel-comments.png)

Les commentaires d’un classeur sont suivis `Workbook.comments` par la propriété. Cela inclut les commentaires créés par les utilisateurs ainsi que les commentaires créés par votre complément. La propriété `Workbook.comments` est un objet [CommentCollection](/javascript/api/excel/excel.commentcollection) qui contient une collection d’objets [Comment](/javascript/api/excel/excel.comment). Les commentaires sont également accessibles au niveau de la [feuille de calcul](/javascript/api/excel/excel.worksheet) . Les exemples de cet article utilisent des commentaires au niveau du classeur, mais ils peuvent être facilement modifiés pour utiliser `Worksheet.comments` la propriété.

## <a name="add-comments"></a>Ajouter des commentaires

Utilisez la `CommentCollection.add` méthode pour ajouter des commentaires à un classeur. Cette méthode peut prendre jusqu’à trois paramètres :

- `cellAddress`: La cellule dans laquelle le commentaire est ajouté. Il peut s’agir d’un objet String ou [Range](/javascript/api/excel/excel.range) . La plage doit être une seule cellule.
- `content`: Contenu du commentaire. Utilisez une chaîne pour les commentaires en texte brut. Utilisez un objet [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) pour les commentaires avec des [mentions](#mentions-online-only). 
- `contentType`: Énumération [ContentType](/javascript/api/excel/excel.contenttype) spécifiant le type de contenu. La valeur par défaut est `ContentType.plain`.

L’exemple de code suivant ajoute un commentaire à la cellule **A2**.

```js
Excel.run(function (context) {
    // Add a comment to A2 on the "MyWorksheet" worksheet.
    var comments = context.workbook.comments;

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `Comment.add`.
    comments.add("MyWorksheet!A2", "TODO: add data.");
    return context.sync();
});
```

> [!NOTE]
> Les commentaires ajoutés par un complément sont attribués à l’utilisateur actuel de ce complément.

### <a name="add-comment-replies"></a>Ajouter des réponses aux commentaires

Un `Comment` objet est un thème de commentaire qui contient zéro ou plusieurs réponses. Les objets `Comment` ont une propriété `replies`, qui est une collection [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) contenant des objets [CommentReply](/javascript/api/excel/excel.commentreply). Pour ajouter une réponse à un commentaire, utilisez la méthode `CommentReplyCollection.add`, en l’appliquant au texte de la réponse. Les réponses s’affichent dans l’ordre dans lequel elles sont ajoutées. Elles sont également attribuées à l’utilisateur actuel du complément.

L’exemple de code suivant ajoute une réponse au premier commentaire du classeur.

```js
Excel.run(function (context) {
    // Get the first comment added to the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    return context.sync();
});
```

## <a name="edit-comments"></a>Modifier les commentaires

Pour modifier un commentaire ou une réponse à un commentaire, configurez sa propriété `Comment.content` ou `CommentReply.content`.

```js
Excel.run(function (context) {
    // Edit the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.content = "PLEASE add headers here.";
    return context.sync();
});
```

### <a name="edit-comment-replies"></a>Modifier les réponses de commentaire

Pour modifier une réponse de commentaire, définissez `CommentReply.content` sa propriété.

```js
Excel.run(function (context) {
    // Edit the first comment reply on the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    var reply = comment.replies.getItemAt(0);
    reply.content = "Never mind";
    return context.sync();
});
```

## <a name="delete-comments"></a>Supprimer les commentaires

Pour supprimer un commentaire, utilisez `Comment.delete` la méthode. La suppression d’un commentaire supprime également les réponses associées à ce commentaire.

```js
Excel.run(function (context) {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    return context.sync();
});
```

### <a name="delete-comment-replies"></a>Supprimer les réponses de commentaire

Pour supprimer une réponse de commentaire, utilisez `CommentReply.delete` la méthode.

```js
Excel.run(function (context) {
    // Delete the first comment reply from this worksheet's first comment.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="resolve-comment-threads-preview"></a>Résoudre les thèmes de commentaire ([Aperçu](../reference/requirement-sets/excel-preview-apis.md)) 

Un thread de commentaire a une valeur `resolved`booléenne configurable, pour indiquer s’il est résolu. Une valeur de `true` signifie que le thread de commentaire est résolu. Une valeur de `false` signifie que le fil de commentaires est nouveau ou rouvert.

```js
Excel.run(function (context) {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    return context.sync();
});
```

Les réponses de commentaire ont `resolved` une propriété ReadOnly. Sa valeur est toujours égale à celle du reste du thread.

## <a name="comment-metadata"></a>Métadonnées de commentaire

Chaque commentaire contient des métadonnées concernant sa création, notamment l’auteur et la date de création. Les commentaires créés par votre complément sont considérés comme créés par l’utilisateur actuel.

L’exemple suivant montre comment afficher l’adresse e-mail et le nom de l’auteur, ainsi que la date de création d’un commentaire dans la cellule **A2**.

```js
Excel.run(function (context) {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    var comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");

    // Load and print the following values.
    comment.load(["authorEmail", "authorName", "creationDate"]);
    return context.sync().then(function () {
        console.log(`${comment.creationDate.toDateString()}: ${comment.authorName} (${comment.authorEmail})`);
    });
});
```

### <a name="comment-reply-metadata"></a>Métadonnées de réponse de commentaire

Les réponses aux commentaires stockent les mêmes types de métadonnées que le commentaire initial.

L’exemple suivant montre comment afficher le courrier électronique, le nom de l’auteur et la date de création de l’auteur de la réponse de commentaire la plus récente à la version **a2**.

```js
Excel.run(function (context) {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    var comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");
    var replyCount = comment.replies.getCount();
    // Sync to get the current number of comment replies.
    return context.sync().then(function () {
        // Get the last comment reply in the comment thread.
        var reply = comment.replies.getItemAt(replyCount.value - 1);
        reply.load(["authorEmail", "authorName", "creationDate"]);
        // Sync to load the reply metadata to print.
        return context.sync().then(function () {
            console.log(`Latest reply: ${reply.creationDate.toDateString()}: ${reply.authorName} ${reply.authorEmail})`);
            return context.sync();
        });
    });
});
```

## <a name="mentions-online-only"></a>Mentions ([en ligne uniquement](../reference/requirement-sets/excel-api-online-requirement-set.md)) 

> [!NOTE]
> Le commentaire mentionne les API sont actuellement disponibles uniquement en préversion publique. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

> [!IMPORTANT]
> Les mentions de commentaire sont actuellement uniquement prises en charge pour Excel sur le Web.

Les [mentions](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd) sont utilisées pour marquer les collègues dans un commentaire. Les notifications sont envoyées avec le contenu de votre commentaire. Votre complément peut créer ces mentions à votre place.

Les commentaires avec des mentions doivent être créés avec des objets [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) . Appelez `CommentCollection.add` avec un `CommentRichContent` conteneur contenant une ou plusieurs mentions et `ContentType.mention` spécifiez `contentType` comme paramètre. La `content` chaîne doit également être mise en forme pour insérer la mention dans le texte. Le format d’une mention est le `<at id="{replyIndex}">{mentionName}</at>`suivant :.

> Note Actuellement, seul le nom exact de la mention peut être utilisé comme texte du lien mention. La prise en charge des versions raccourcies d’un nom sera ajoutée ultérieurement.

L’exemple suivant montre un commentaire avec une seule mention.

```js
Excel.run(function (context) {
    // Add an "@mention" for "Kate Kristensen" to cell A1 in the "MyWorksheet" worksheet.
    var mention = {
        email: "kakri@contoso.com",
        id: 0,
        name: "Kate Kristensen"
    };

    // This will tag the mention's name using the '@' syntax.
    // They will be notified via email.
    var commentBody = {
        mentions: [mention],
        richContent: '<at id="0">' + mention.name + "</at> -  Can you take a look?"
    };

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `comment.add`.
    context.workbook.comments.add("MyWorksheet!A1", commentBody, Excel.ContentType.mention);
    return context.sync();
});
```

## <a name="see-also"></a>Voir aussi

- [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](excel-add-ins-core-concepts.md)
- [Utiliser les classeurs utilisant l’API JavaScript Excel](excel-add-ins-workbooks.md)
- [Insérer des commentaires et des notes dans Excel](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
