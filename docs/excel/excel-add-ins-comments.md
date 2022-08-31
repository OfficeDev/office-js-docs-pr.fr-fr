---
title: Utiliser des commentaires à l’aide de l’API JavaScript Excel
description: Informations sur l’utilisation des API pour ajouter, supprimer et modifier des commentaires et des threads de commentaires.
ms.date: 02/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5996c1bb55c3d4a358786b15f7c3e46aae6f42aa
ms.sourcegitcommit: eef2064d7966db91f8401372dd255a32d76168c2
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/31/2022
ms.locfileid: "67464796"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a>Utiliser des commentaires à l’aide de l’API JavaScript Excel

Cet article explique comment ajouter, lire, modifier et supprimer des commentaires dans un classeur avec l’API JavaScript Excel. Pour en savoir plus sur la fonctionnalité de commentaire, consultez l’article [Insérer des commentaires et des notes dans Excel](https://support.microsoft.com/office/bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) .

Dans l’API JavaScript Excel, un commentaire inclut à la fois le commentaire initial unique et la discussion thématique connectée. Il est lié à une cellule individuelle. Toute personne qui affiche le classeur avec des autorisations suffisantes peut répondre à un commentaire. Un objet [Comment](/javascript/api/excel/excel.comment) stocke ces réponses en tant qu’objets [CommentReply](/javascript/api/excel/excel.commentreply) . Vous devez considérer un commentaire comme un thread et qu’un thread doit avoir une entrée spéciale comme point de départ.

![Un commentaire Excel, intitulé « Commentaire » avec deux réponses, étiqueté « Comment.replies[0] » et « Comment.replies[1].](../images/excel-comments.png)

Les commentaires d’un classeur sont suivis par la `Workbook.comments` propriété. Cela inclut les commentaires créés par les utilisateurs ainsi que les commentaires créés par votre complément. La propriété `Workbook.comments` est un objet [CommentCollection](/javascript/api/excel/excel.commentcollection) qui contient une collection d’objets [Comment](/javascript/api/excel/excel.comment). Les commentaires sont également accessibles au niveau [de la feuille de calcul](/javascript/api/excel/excel.worksheet) . Les exemples de cet article fonctionnent avec des commentaires au niveau du classeur, mais ils peuvent être facilement modifiés pour utiliser la `Worksheet.comments` propriété.

## <a name="add-comments"></a>Ajouter des commentaires

Utilisez la `CommentCollection.add` méthode pour ajouter des commentaires à un classeur. Cette méthode prend jusqu’à trois paramètres :

- `cellAddress`: cellule dans laquelle le commentaire est ajouté. Il peut s’agir d’une chaîne ou [d’un objet Range](/javascript/api/excel/excel.range) . La plage doit être une cellule unique.
- `content`: contenu du commentaire. Utilisez une chaîne pour les commentaires en texte brut. Utilisez un objet [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) pour les commentaires avec [mentions](#mentions).
- `contentType`: énumération [ContentType](/javascript/api/excel/excel.contenttype) spécifiant le type de contenu. La valeur par défaut est `ContentType.plain`.

L’exemple de code suivant ajoute un commentaire à la cellule **A2**.

```js
await Excel.run(async (context) => {
    // Add a comment to A2 on the "MyWorksheet" worksheet.
    let comments = context.workbook.comments;

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `Comment.add`.
    comments.add("MyWorksheet!A2", "TODO: add data.");
    await context.sync();
});
```

> [!NOTE]
> Les commentaires ajoutés par un complément sont attribués à l’utilisateur actuel de ce complément.

### <a name="add-comment-replies"></a>Ajouter des réponses de commentaire

Un `Comment` objet est un thread de commentaire qui contient zéro ou plusieurs réponses. Les objets `Comment` ont une propriété `replies`, qui est une collection [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) contenant des objets [CommentReply](/javascript/api/excel/excel.commentreply). Pour ajouter une réponse à un commentaire, utilisez la méthode `CommentReplyCollection.add`, en l’appliquant au texte de la réponse. Les réponses s’affichent dans l’ordre dans lequel elles sont ajoutées. Ils sont également attribués à l’utilisateur actuel du complément.

L’exemple de code suivant ajoute une réponse au premier commentaire du classeur.

```js
await Excel.run(async (context) => {
    // Get the first comment added to the workbook.
    let comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    await context.sync();
});
```

## <a name="edit-comments"></a>Modifier les commentaires

Pour modifier un commentaire ou une réponse à un commentaire, configurez sa propriété `Comment.content` ou `CommentReply.content`.

```js
await Excel.run(async (context) => {
    // Edit the first comment in the workbook.
    let comment = context.workbook.comments.getItemAt(0);
    comment.content = "PLEASE add headers here.";
    await context.sync();
});
```

### <a name="edit-comment-replies"></a>Modifier les réponses aux commentaires

Pour modifier une réponse de commentaire, définissez sa `CommentReply.content` propriété.

```js
await Excel.run(async (context) => {
    // Edit the first comment reply on the first comment in the workbook.
    let comment = context.workbook.comments.getItemAt(0);
    let reply = comment.replies.getItemAt(0);
    reply.content = "Never mind";
    await context.sync();
});
```

## <a name="delete-comments"></a>Supprimer les commentaires

Pour supprimer un commentaire, utilisez la `Comment.delete` méthode. La suppression d’un commentaire supprime également les réponses associées à ce commentaire.

```js
await Excel.run(async (context) => {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    await context.sync();
});
```

### <a name="delete-comment-replies"></a>Supprimer les réponses de commentaire

Pour supprimer une réponse de commentaire, utilisez la `CommentReply.delete` méthode.

```js
await Excel.run(async (context) => {
    // Delete the first comment reply from this worksheet's first comment.
    let comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    await context.sync();
});
```

## <a name="resolve-comment-threads"></a>Résoudre les threads de commentaire

Un thread de commentaire a une valeur booléenne configurable, `resolved`pour indiquer s’il est résolu. La valeur signifie que le thread de `true` commentaire est résolu. La valeur signifie que le thread de `false` commentaire est nouveau ou rouvert.

```js
await Excel.run(async (context) => {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    await context.sync();
});
```

Les réponses aux commentaires ont une propriété en lecture seule `resolved` . Sa valeur est toujours égale à celle du reste du thread.

## <a name="comment-metadata"></a>Métadonnées de commentaire

Chaque commentaire contient des métadonnées concernant sa création, notamment l’auteur et la date de création. Les commentaires créés par votre complément sont considérés comme créés par l’utilisateur actuel.

L’exemple suivant montre comment afficher l’adresse e-mail et le nom de l’auteur, ainsi que la date de création d’un commentaire dans la cellule **A2**.

```js
await Excel.run(async (context) => {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    let comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");

    // Load and print the following values.
    comment.load(["authorEmail", "authorName", "creationDate"]);
    await context.sync();
    
    console.log(`${comment.creationDate.toDateString()}: ${comment.authorName} (${comment.authorEmail})`);
});
```

### <a name="comment-reply-metadata"></a>Métadonnées de réponse de commentaire

Les réponses de commentaire stockent les mêmes types de métadonnées que le commentaire initial.

L’exemple suivant montre comment afficher l’e-mail de l’auteur, le nom de l’auteur et la date de création de la dernière réponse de commentaire sur **A2**.

```js
await Excel.run(async (context) => {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    let comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");
    let replyCount = comment.replies.getCount();
    // Sync to get the current number of comment replies.
    await context.sync();

    // Get the last comment reply in the comment thread.
    let reply = comment.replies.getItemAt(replyCount.value - 1);
    reply.load(["authorEmail", "authorName", "creationDate"]);

    // Sync to load the reply metadata to print.
    await context.sync();

    console.log(`Latest reply: ${reply.creationDate.toDateString()}: ${reply.authorName} ${reply.authorEmail})`);
    await context.sync();
});
```

## <a name="mentions"></a>Mentions

[Les mentions sont utilisées](https://support.microsoft.com/office/644bf689-31a0-4977-a4fb-afe01820c1fd) pour marquer des collègues dans un commentaire. Cela leur envoie des notifications avec le contenu de votre commentaire. Votre complément peut créer ces mentions en votre nom.

Les commentaires avec mentions doivent être créés avec [des objets CommentRichContent](/javascript/api/excel/excel.commentrichcontent) . Appelez `CommentCollection.add` avec une `CommentRichContent` ou plusieurs mentions contenantes et spécifiez `ContentType.mention` comme `contentType` paramètre. La `content` chaîne doit également être mise en forme pour insérer la mention dans le texte. Le format d’une mention est : `<at id="{replyIndex}">{mentionName}</at>`.

> [!NOTE]
> Actuellement, seul le nom exact de la mention peut être utilisé comme texte du lien de mention. La prise en charge des versions abrégées d’un nom sera ajoutée ultérieurement.

L’exemple suivant montre un commentaire avec une seule mention.

```js
await Excel.run(async (context) => {
    // Add an "@mention" for "Kate Kristensen" to cell A1 in the "MyWorksheet" worksheet.
    let mention = {
        email: "kakri@contoso.com",
        id: 0,
        name: "Kate Kristensen"
    };

    // This will tag the mention's name using the '@' syntax.
    // They will be notified via email.
    let commentBody = {
        mentions: [mention],
        richContent: '<at id="0">' + mention.name + "</at> -  Can you take a look?"
    };

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `comment.add`.
    context.workbook.comments.add("MyWorksheet!A1", commentBody, Excel.ContentType.mention);
    await context.sync();
});
```

## <a name="comment-events"></a>Événements de commentaire

Votre complément peut écouter les ajouts de commentaires, les modifications et les suppressions. [Les événements de commentaire](/javascript/api/excel/excel.commentcollection#event-details) se produisent sur l’objet `CommentCollection` . Pour écouter les événements de commentaire, inscrivez le `onAdded`gestionnaire d’événements , `onChanged`ou `onDeleted` comment. Lorsqu’un événement de commentaire est détecté, utilisez ce gestionnaire d’événements pour récupérer des données sur le commentaire ajouté, modifié ou supprimé. L’événement `onChanged` gère également les ajouts, modifications et suppressions de réponse aux commentaires.

Chaque événement de commentaire se déclenche une seule fois lorsque plusieurs ajouts, modifications ou suppressions sont effectués en même temps. Tous les objets [CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs), [CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs) et [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) contiennent des tableaux d’ID de commentaire pour mapper les actions d’événement aux collections de commentaires.

Pour plus d’informations sur l’inscription des gestionnaires d’événements, la gestion des événements et la suppression des gestionnaires d’événements, consultez l’article [Travailler avec les événements à l’aide de l’API JavaScript Excel](excel-add-ins-events.md) .

### <a name="comment-addition-events"></a>Commentaires sur les événements d’ajout

L’événement `onAdded` est déclenché lorsqu’un ou plusieurs nouveaux commentaires sont ajoutés à la collection de commentaires. Cet événement *n’est pas* déclenché lorsque des réponses sont ajoutées à un thread de commentaire (consultez [les événements de modification](#comment-change-events) de commentaire pour en savoir plus sur les événements de réponse de commentaire).

L’exemple suivant montre comment inscrire le `onAdded` gestionnaire d’événements, puis utiliser l’objet `CommentAddedEventArgs` pour récupérer le `commentDetails` tableau du commentaire ajouté.

> [!NOTE]
> Cet exemple fonctionne uniquement lorsqu’un seul commentaire est ajouté.

```js
await Excel.run(async (context) => {
    let comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onAdded comment event handler.
    comments.onAdded.add(commentAdded);

    await context.sync();
});

async function commentAdded() {
    await Excel.run(async (context) => {
        // Retrieve the added comment using the comment ID.
        // Note: This method assumes only a single comment is added at a time. 
        let addedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);

        // Load the added comment's data.
        addedComment.load(["content", "authorName"]);

        await context.sync();

        // Print out the added comment's data.
        console.log(`A comment was added. ID: ${event.commentDetails[0].commentId}. Comment content:${addedComment.content}. Comment author:${addedComment.authorName}`);
        await context.sync();
    });
}
```

### <a name="comment-change-events"></a>Commenter les événements de modification

L’événement `onChanged` de commentaire est déclenché dans les scénarios suivants.

- Le contenu d’un commentaire est mis à jour.
- Un thread de commentaire est résolu.
- Un thread de commentaire est rouvert.
- Une réponse est ajoutée à un thread de commentaire.
- Une réponse est mise à jour dans un thread de commentaire.
- Une réponse est supprimée dans un thread de commentaire.

L’exemple suivant montre comment inscrire le `onChanged` gestionnaire d’événements, puis utiliser l’objet `CommentChangedEventArgs` pour récupérer le `commentDetails` tableau du commentaire modifié.

> [!NOTE]
> Cet exemple fonctionne uniquement lorsqu’un seul commentaire est modifié.

```js
await Excel.run(async (context) => {
    let comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onChanged comment event handler.
    comments.onChanged.add(commentChanged);

    await context.sync();
});

async function commentChanged() {
    await Excel.run(async (context) => {
        // Retrieve the changed comment using the comment ID.
        // Note: This method assumes only a single comment is changed at a time. 
        let changedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);

        // Load the changed comment's data.
        changedComment.load(["content", "authorName"]);

        await context.sync();

        // Print out the changed comment's data.
        console.log(`A comment was changed. ID: ${event.commentDetails[0].commentId}. Updated comment content: ${changedComment.content}. Comment author: ${changedComment.authorName}`);
        await context.sync();
    });
}
```

### <a name="comment-deletion-events"></a>Commenter les événements de suppression

L’événement `onDeleted` est déclenché lorsqu’un commentaire est supprimé de la collection de commentaires. Une fois qu’un commentaire a été supprimé, ses métadonnées ne sont plus disponibles. L’objet [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) fournit des ID de commentaire, au cas où votre complément gérerait des commentaires individuels.

L’exemple suivant montre comment inscrire le `onDeleted` gestionnaire d’événements, puis utiliser l’objet `CommentDeletedEventArgs` pour récupérer le `commentDetails` tableau du commentaire supprimé.

> [!NOTE]
> Cet exemple fonctionne uniquement lorsqu’un seul commentaire est supprimé.

```js
await Excel.run(async (context) => {
    let comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onDeleted comment event handler.
    comments.onDeleted.add(commentDeleted);

    await context.sync();
});

async function commentDeleted() {
    await Excel.run(async (context) => {
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
