---
title: Utiliser des commentaires à l’aide de l’API JavaScript pour Excel
description: Informations sur l’utilisation des API pour ajouter, supprimer et modifier des commentaires et des thèmes de commentaires.
ms.date: 10/09/2020
localization_priority: Normal
ms.openlocfilehash: 00f7dd22fb2148902152197521098482071e5284
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626420"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a>Utiliser des commentaires à l’aide de l’API JavaScript pour Excel

Cet article explique comment ajouter, lire, modifier et supprimer des commentaires dans un classeur à l’aide de l’API JavaScript pour Excel. Pour en savoir plus sur la fonctionnalité de commentaire, consultez l’article [Insérer des commentaires et des notes dans Excel](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) .

Dans l’API JavaScript pour Excel, un commentaire inclut à la fois le commentaire initial unique et la discussion liée au thread. Elle est liée à une cellule individuelle. Toute personne qui consulte le classeur avec des autorisations suffisantes peut répondre à un commentaire. Un objet [Comment](/javascript/api/excel/excel.comment) stocke ces réponses en tant qu’objets [CommentReply](/javascript/api/excel/excel.commentreply) . Vous devez considérer un commentaire comme un fil de discussion et qu’un thread doit avoir une entrée spéciale comme point de départ.

![Commentaire Excel, étiqueté « commentaire » avec deux réponses, intitulées « comment. réponses [0] » et «comment. réponses [1].](../images/excel-comments.png)

Les commentaires d’un classeur sont suivis par la `Workbook.comments` propriété. Cela inclut les commentaires créés par les utilisateurs ainsi que les commentaires créés par votre complément. La propriété `Workbook.comments` est un objet [CommentCollection](/javascript/api/excel/excel.commentcollection) qui contient une collection d’objets [Comment](/javascript/api/excel/excel.comment). Les commentaires sont également accessibles au niveau de la [feuille de calcul](/javascript/api/excel/excel.worksheet) . Les exemples de cet article utilisent des commentaires au niveau du classeur, mais ils peuvent être facilement modifiés pour utiliser la `Worksheet.comments` propriété.

## <a name="add-comments"></a>Ajouter des commentaires

Utilisez la `CommentCollection.add` méthode pour ajouter des commentaires à un classeur. Cette méthode peut prendre jusqu’à trois paramètres :

- `cellAddress`: La cellule dans laquelle le commentaire est ajouté. Il peut s’agir d’un objet String ou [Range](/javascript/api/excel/excel.range) . La plage doit être une seule cellule.
- `content`: Contenu du commentaire. Utilisez une chaîne pour les commentaires en texte brut. Utilisez un objet [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) pour les commentaires avec des [mentions](#mentions).
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

Pour modifier une réponse de commentaire, définissez sa `CommentReply.content` propriété.

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

Pour supprimer un commentaire, utilisez la `Comment.delete` méthode. La suppression d’un commentaire supprime également les réponses associées à ce commentaire.

```js
Excel.run(function (context) {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    return context.sync();
});
```

### <a name="delete-comment-replies"></a>Supprimer les réponses de commentaire

Pour supprimer une réponse de commentaire, utilisez la `CommentReply.delete` méthode.

```js
Excel.run(function (context) {
    // Delete the first comment reply from this worksheet's first comment.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="resolve-comment-threads"></a>Résoudre les thèmes de commentaires

Un thread de commentaire a une valeur booléenne configurable, `resolved` pour indiquer s’il est résolu. Une valeur de `true` signifie que le thread de commentaire est résolu. Une valeur de `false` signifie que le fil de commentaires est nouveau ou rouvert.

```js
Excel.run(function (context) {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    return context.sync();
});
```

Les réponses de commentaire ont une `resolved` propriété ReadOnly. Sa valeur est toujours égale à celle du reste du thread.

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

## <a name="mentions"></a>Mentions

Les [mentions](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd) sont utilisées pour marquer les collègues dans un commentaire. Les notifications sont envoyées avec le contenu de votre commentaire. Votre complément peut créer ces mentions à votre place.

Les commentaires avec des mentions doivent être créés avec des objets [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) . Appelez `CommentCollection.add` avec un `CommentRichContent` conteneur contenant une ou plusieurs mentions et spécifiez `ContentType.mention` comme `contentType` paramètre. La `content` chaîne doit également être mise en forme pour insérer la mention dans le texte. Le format d’une mention est le suivant : `<at id="{replyIndex}">{mentionName}</at>` .

> [!NOTE]
> Actuellement, seul le nom exact de la mention peut être utilisé comme texte du lien mention. La prise en charge des versions raccourcies d’un nom sera ajoutée ultérieurement.

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

## <a name="comment-events"></a>Événements de commentaire

Votre complément peut écouter les ajouts, les modifications et les suppressions de commentaires. Les [événements de commentaire](/javascript/api/excel/excel.commentcollection#event-details) se produisent sur l' `CommentCollection` objet. Pour écouter les événements de commentaire, enregistrez `onAdded` le `onChanged` Gestionnaire d’événements,, ou le `onDeleted` commentaire. Lorsqu’un événement de commentaire est détecté, utilisez ce gestionnaire d’événements pour récupérer des données sur le Commentaire ajouté, modifié ou supprimé. L' `onChanged` événement gère également les ajouts de réponse aux commentaires, les modifications et les suppressions. 

Chaque événement de commentaire ne déclenche qu’une seule fois lorsque plusieurs ajouts, modifications ou suppressions sont effectués en même temps. Tous les objets [CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs), [CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)et [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) contiennent des tableaux d’ID de commentaires permettant de mapper les actions d’événement vers les collections de commentaires.

Pour plus d’informations sur l’inscription de gestionnaires d’événements, la gestion des événements et la suppression de gestionnaires d’événements, voir l’article [work with Events using the Excel JavaScript API](excel-add-ins-events.md) . 

### <a name="comment-addition-events"></a>Événements d’ajout de commentaires 
L' `onAdded` événement est déclenché lorsqu’un ou plusieurs nouveaux commentaires sont ajoutés à la collection de commentaires. Cet événement n’est *pas* déclenché lorsque les réponses sont ajoutées à un thread de commentaire (voir [événements de modification](#comment-change-events) des commentaires pour en savoir plus sur les événements de réponse aux commentaires).

L’exemple suivant montre comment inscrire le `onAdded` Gestionnaire d’événements, puis utiliser l' `CommentAddedEventArgs` objet pour récupérer le `commentDetails` tableau du Commentaire ajouté.

> [!NOTE]
> Cet exemple fonctionne uniquement lorsqu’un seul commentaire est ajouté. 

```js
Excel.run(function (context) {
    var comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onAdded comment event handler.
    comments.onAdded.add(commentAdded);

    return context.sync();
});

function commentAdded() {
    Excel.run(function (context) {
        // Retrieve the added comment using the comment ID.
        // Note: This method assumes only a single comment is added at a time. 
        var addedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);

        // Load the added comment's data.
        addedComment.load(["content", "authorName"]);

        return context.sync().then(function () {
            // Print out the added comment's data.
            console.log(`A comment was added. ID: ${event.commentDetails[0].commentId}. Comment content:${addedComment.content}. Comment author:${addedComment.authorName}`);
            return context.sync();
        });            
    });
}
```

### <a name="comment-change-events"></a>Événements de modification de commentaire 
L' `onChanged` événement comment est déclenché dans les scénarios suivants.

- Le contenu d’un commentaire est mis à jour.
- Une thread de commentaire est résolue.
- Une thread de commentaire est rouverte.
- Une réponse est ajoutée à une thread de commentaire.
- Une réponse est mise à jour dans une thread de commentaire.
- Une réponse est supprimée dans une thread de commentaire.

L’exemple suivant montre comment inscrire le `onChanged` Gestionnaire d’événements, puis utiliser l' `CommentChangedEventArgs` objet pour récupérer le `commentDetails` tableau du commentaire modifié.

> [!NOTE]
> Cet exemple fonctionne uniquement lorsqu’un seul commentaire est modifié. 

```js
Excel.run(function (context) {
    var comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onChanged comment event handler.
    comments.onChanged.add(commentChanged);

    return context.sync();
});    

function commentChanged() {
    Excel.run(function (context) {
        // Retrieve the changed comment using the comment ID.
        // Note: This method assumes only a single comment is changed at a time. 
        var changedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);

        // Load the changed comment's data.
        changedComment.load(["content", "authorName"]);

        return context.sync().then(function () {
            // Print out the changed comment's data.
            console.log(`A comment was changed. ID: ${event.commentDetails[0].commentId}`. Updated comment content: ${changedComment.content}`. Comment author: ${changedComment.authorName}`);
            return context.sync();
        });
    });
}
```

### <a name="comment-deletion-events"></a>Événements de suppression de commentaires
L' `onDeleted` événement est déclenché lorsqu’un commentaire est supprimé de la collection de commentaires. Une fois qu’un commentaire a été supprimé, ses métadonnées ne sont plus disponibles. L’objet [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) fournit des ID de commentaire, si votre complément gère des Commentaires individuels.

L’exemple suivant montre comment inscrire le `onDeleted` Gestionnaire d’événements, puis utiliser l' `CommentDeletedEventArgs` objet pour récupérer le `commentDetails` tableau du commentaire supprimé.

> [!NOTE]
> Cet exemple ne fonctionne qu’en cas de suppression d’un seul commentaire. 

```js
Excel.run(function (context) {
    var comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onDeleted comment event handler.
    comments.onDeleted.add(commentDeleted);

    return context.sync();
});

function commentDeleted() {
    Excel.run(function (context) {
        // Print out the deleted comment's ID.
        // Note: This method assumes only a single comment is deleted at a time. 
        console.log(`A comment was deleted. ID: ${event.commentDetails[0].commentId}`);
    });
}
```

## <a name="see-also"></a>Voir aussi

- [Modèle objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Utiliser les classeurs utilisant l’API JavaScript Excel](excel-add-ins-workbooks.md)
- [Utilisation d’événements à l’aide de l’API JavaScript pour Excel](excel-add-ins-events.md)
- [Insérer des commentaires et des notes dans Excel](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
