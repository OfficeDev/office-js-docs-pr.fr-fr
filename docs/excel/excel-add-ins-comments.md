---
title: Utiliser des commentaires à l’aide de l Excel API JavaScript
description: Informations sur l’utilisation des API pour ajouter, supprimer et modifier des commentaires et des threads de commentaires.
ms.date: 10/09/2020
localization_priority: Normal
ms.openlocfilehash: 16569bc1d72391dff0ac35a48e45470ff90852f8
ms.sourcegitcommit: 69f6492de8a4c91e734250c76681c44b3f349440
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/03/2021
ms.locfileid: "58868651"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a>Utiliser des commentaires à l’aide de l Excel API JavaScript

Cet article explique comment ajouter, lire, modifier et supprimer des commentaires dans un Excel api JavaScript. Pour en savoir plus sur la fonctionnalité de commentaire, voir l’article Insérer des commentaires et des [notes Excel’article.](https://support.microsoft.com/office/bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)

Dans l Excel API JavaScript, un commentaire inclut à la fois le commentaire initial unique et la discussion threaded connectée. Elle est liée à une cellule individuelle. Toute personne qui affiche le manuel avec des autorisations suffisantes peut répondre à un commentaire. Un [objet Comment](/javascript/api/excel/excel.comment) stocke ces réponses en tant [qu’objets CommentReply.](/javascript/api/excel/excel.commentreply) Vous devez considérer un commentaire comme un thread et qu’un thread doit avoir une entrée spéciale comme point de départ.

![Un Excel, étiqueté « Comment » avec deux réponses, étiqueté « Comment.replies[0] » et « Comment.replies[1].](../images/excel-comments.png)

Les commentaires dans un workbook sont suivis par la `Workbook.comments` propriété. Cela inclut les commentaires créés par les utilisateurs ainsi que les commentaires créés par votre complément. La propriété `Workbook.comments` est un objet [CommentCollection](/javascript/api/excel/excel.commentcollection) qui contient une collection d’objets [Comment](/javascript/api/excel/excel.comment). Les commentaires sont également accessibles au niveau [de la feuille de](/javascript/api/excel/excel.worksheet) calcul. Les exemples de cet article fonctionnent avec des commentaires au niveau du workbook, mais ils peuvent être facilement modifiés pour utiliser la `Worksheet.comments` propriété.

## <a name="add-comments"></a>Ajouter des commentaires

Utilisez la `CommentCollection.add` méthode pour ajouter des commentaires à un workbook. Cette méthode prend jusqu’à trois paramètres :

- `cellAddress`: cellule dans laquelle le commentaire est ajouté. Il peut s’agit d’une chaîne ou [d’un objet Range.](/javascript/api/excel/excel.range) La plage doit être une cellule unique.
- `content`: contenu du commentaire. Utilisez une chaîne pour les commentaires en texte simple. Utilisez un [objet CommentRichContent](/javascript/api/excel/excel.commentrichcontent) pour les commentaires avec [des mentions.](#mentions)
- `contentType`: Une [enum ContentType](/javascript/api/excel/excel.contenttype) spécifiant le type de contenu. La valeur par défaut est `ContentType.plain`.

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
> Les commentaires ajoutés par un add-in sont attribués à l’utilisateur actuel de ce dernier.

### <a name="add-comment-replies"></a>Ajouter des réponses de commentaire

Un `Comment` objet est un thread de commentaires qui contient zéro ou plusieurs réponses. Les objets `Comment` ont une propriété `replies`, qui est une collection [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) contenant des objets [CommentReply](/javascript/api/excel/excel.commentreply). Pour ajouter une réponse à un commentaire, utilisez la méthode `CommentReplyCollection.add`, en l’appliquant au texte de la réponse. Les réponses s’affichent dans l’ordre dans lequel elles sont ajoutées. Ils sont également attribués à l’utilisateur actuel du module.

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

### <a name="edit-comment-replies"></a>Modifier les réponses aux commentaires

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

### <a name="delete-comment-replies"></a>Supprimer les réponses aux commentaires

Pour supprimer une réponse de commentaire, utilisez la `CommentReply.delete` méthode.

```js
Excel.run(function (context) {
    // Delete the first comment reply from this worksheet's first comment.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="resolve-comment-threads"></a>Résoudre les threads de commentaires

Un thread de commentaire a une valeur boolénable configurable, pour indiquer `resolved` si elle est résolue. Une valeur de `true` signifie que le thread de commentaire est résolu. Une valeur de `false` signifie que le thread de commentaire est nouveau ou rouvert.

```js
Excel.run(function (context) {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    return context.sync();
});
```

Les réponses de commentaire ont une propriété `resolved` readonly. Sa valeur est toujours égale à celle du reste du thread.

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

### <a name="comment-reply-metadata"></a>Commenter les métadonnées de réponse

Les réponses de commentaire stockent les mêmes types de métadonnées que le commentaire initial.

L’exemple suivant montre comment afficher le courrier électronique de l’auteur, le nom de l’auteur et la date de création de la dernière réponse de commentaire **sur A2**.

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

[Les mentions](https://support.microsoft.com/office/644bf689-31a0-4977-a4fb-afe01820c1fd) sont utilisées pour marquer des collègues dans un commentaire. Cela leur envoie des notifications avec le contenu de votre commentaire. Votre add-in peut créer ces mentions en votre nom.

Les commentaires avec mentions doivent être créés avec des objets [CommentRichContent.](/javascript/api/excel/excel.commentrichcontent) Appelez `CommentCollection.add` avec une ou plusieurs `CommentRichContent` mentions contenantes et `ContentType.mention` spécifiez en tant que `contentType` paramètre. La `content` chaîne doit également être mise en forme pour insérer la mention dans le texte. Le format d’une mention est : `<at id="{replyIndex}">{mentionName}</at>` .

> [!NOTE]
> Actuellement, seul le nom exact de la mention peut être utilisé comme texte du lien de mention. La prise en charge des versions raccourcies d’un nom sera ajoutée ultérieurement.

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

Votre complément peut écouter les ajouts de commentaires, les modifications et les suppressions. [Les événements de commentaire](/javascript/api/excel/excel.commentcollection#event-details) se produisent sur `CommentCollection` l’objet. Pour écouter les événements de commentaire, inscrivez le `onAdded` `onChanged` , ou le `onDeleted` handler d’événements de commentaire. Lorsqu’un événement de commentaire est détecté, utilisez ce handler d’événements pour récupérer des données sur le commentaire ajouté, modifié ou supprimé. L’événement gère également les ajouts, modifications et suppressions de réponses `onChanged` aux commentaires. 

Chaque événement de commentaire se déclenche une seule fois lorsque plusieurs ajouts, modifications ou suppressions sont effectués en même temps. Tous les objets [CommentAddedEventArgs,](/javascript/api/excel/excel.commentaddedeventargs) [CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)et [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) contiennent des tableaux d’ID de commentaire pour maguer les actions d’événement aux collections de commentaires.

Consultez l’article Utiliser des événements à l’aide de [l’API JavaScript Excel](excel-add-ins-events.md) pour plus d’informations sur l’inscription des handlers d’événements, la gestion des événements et la suppression de ces derniers. 

### <a name="comment-addition-events"></a>Événements d’ajout de commentaires 
L’événement est déclenché lorsqu’un ou plusieurs nouveaux commentaires sont `onAdded` ajoutés à la collection de commentaires. Cet événement *n’est pas* déclenché lorsque des réponses sont ajoutées à un thread de commentaires (voir Comment change [events](#comment-change-events) to learn about comment reply events).

L’exemple suivant montre comment inscrire le handler d’événements, puis utiliser l’objet pour récupérer le tableau `onAdded` `CommentAddedEventArgs` du commentaire `commentDetails` ajouté.

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
`onChanged`L’événement de commentaire est déclenché dans les scénarios suivants.

- Le contenu d’un commentaire est mis à jour.
- Un thread de commentaire est résolu.
- Un thread de commentaire est rouvert.
- Une réponse est ajoutée à un thread de commentaires.
- Une réponse est mise à jour dans un thread de commentaires.
- Une réponse est supprimée dans un thread de commentaires.

L’exemple suivant montre comment inscrire le handler d’événements, puis utiliser l’objet pour récupérer le tableau `onChanged` `CommentChangedEventArgs` du commentaire `commentDetails` modifié.

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
`onDeleted`L’événement est déclenché lorsqu’un commentaire est supprimé de la collection de commentaires. Une fois qu’un commentaire a été supprimé, ses métadonnées ne sont plus disponibles. [L’objet CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) fournit des ID de commentaire, au cas où votre add-in gère des commentaires individuels.

L’exemple suivant montre comment inscrire le handler d’événements, puis utiliser l’objet pour récupérer le tableau `onDeleted` `CommentDeletedEventArgs` du commentaire `commentDetails` supprimé.

> [!NOTE]
> Cet exemple fonctionne uniquement lorsqu’un seul commentaire est supprimé. 

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

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Utiliser les classeurs utilisant l’API JavaScript Excel](excel-add-ins-workbooks.md)
- [Utilisation d’événements à l’aide de l’API JavaScript pour Excel](excel-add-ins-events.md)
- [Insérer des commentaires et des notes dans Excel](https://support.microsoft.com/office/bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
