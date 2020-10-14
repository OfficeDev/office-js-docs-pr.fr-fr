---
title: Utiliser des commentaires à l’aide de l’API JavaScript pour Excel
description: Informations sur l’utilisation des API pour ajouter, supprimer et modifier des commentaires et des thèmes de commentaires.
ms.date: 10/09/2020
localization_priority: Normal
ms.openlocfilehash: 85312cbd92aa6c9d0f82fd167e8a372c2eff8c85
ms.sourcegitcommit: b50eebd303adcc22eb86e65756ce7e9a82f41a57
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/14/2020
ms.locfileid: "48456551"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a><span data-ttu-id="291df-103">Utiliser des commentaires à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="291df-103">Work with comments using the Excel JavaScript API</span></span>

<span data-ttu-id="291df-104">Cet article explique comment ajouter, lire, modifier et supprimer des commentaires dans un classeur à l’aide de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="291df-104">This article describes how to add, read, modify, and remove comments in a workbook with the Excel JavaScript API.</span></span> <span data-ttu-id="291df-105">Pour en savoir plus sur la fonctionnalité de commentaire, consultez l’article [Insérer des commentaires et des notes dans Excel](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) .</span><span class="sxs-lookup"><span data-stu-id="291df-105">You can learn more about the comment feature from the [Insert comments and notes in Excel](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) article.</span></span>

<span data-ttu-id="291df-106">Dans l’API JavaScript pour Excel, un commentaire inclut à la fois le commentaire initial unique et la discussion liée au thread.</span><span class="sxs-lookup"><span data-stu-id="291df-106">In the Excel JavaScript API, a comment includes both the single initial comment and the connected threaded discussion.</span></span> <span data-ttu-id="291df-107">Elle est liée à une cellule individuelle.</span><span class="sxs-lookup"><span data-stu-id="291df-107">It is tied to an individual cell.</span></span> <span data-ttu-id="291df-108">Toute personne qui consulte le classeur avec des autorisations suffisantes peut répondre à un commentaire.</span><span class="sxs-lookup"><span data-stu-id="291df-108">Anyone viewing the workbook with sufficient permissions can reply to a comment.</span></span> <span data-ttu-id="291df-109">Un objet [Comment](/javascript/api/excel/excel.comment) stocke ces réponses en tant qu’objets [CommentReply](/javascript/api/excel/excel.commentreply) .</span><span class="sxs-lookup"><span data-stu-id="291df-109">A [Comment](/javascript/api/excel/excel.comment) object stores those replies as [CommentReply](/javascript/api/excel/excel.commentreply) objects.</span></span> <span data-ttu-id="291df-110">Vous devez considérer un commentaire comme un fil de discussion et qu’un thread doit avoir une entrée spéciale comme point de départ.</span><span class="sxs-lookup"><span data-stu-id="291df-110">You should consider a comment to be a thread and that a thread must have a special entry as the starting point.</span></span>

![Commentaire Excel, étiqueté « commentaire » avec deux réponses, intitulées « comment. réponses [0] » et «comment. réponses [1].](../images/excel-comments.png)

<span data-ttu-id="291df-112">Les commentaires d’un classeur sont suivis par la `Workbook.comments` propriété.</span><span class="sxs-lookup"><span data-stu-id="291df-112">Comments within a workbook are tracked by the `Workbook.comments` property.</span></span> <span data-ttu-id="291df-113">Cela inclut les commentaires créés par les utilisateurs ainsi que les commentaires créés par votre complément.</span><span class="sxs-lookup"><span data-stu-id="291df-113">This includes comments created by users and also comments created by your add-in.</span></span> <span data-ttu-id="291df-114">La propriété `Workbook.comments` est un objet [CommentCollection](/javascript/api/excel/excel.commentcollection) qui contient une collection d’objets [Comment](/javascript/api/excel/excel.comment).</span><span class="sxs-lookup"><span data-stu-id="291df-114">The `Workbook.comments` property is a [CommentCollection](/javascript/api/excel/excel.commentcollection) object that contains a collection of [Comment](/javascript/api/excel/excel.comment) objects.</span></span> <span data-ttu-id="291df-115">Les commentaires sont également accessibles au niveau de la [feuille de calcul](/javascript/api/excel/excel.worksheet) .</span><span class="sxs-lookup"><span data-stu-id="291df-115">Comments are also accessible at the [Worksheet](/javascript/api/excel/excel.worksheet) level.</span></span> <span data-ttu-id="291df-116">Les exemples de cet article utilisent des commentaires au niveau du classeur, mais ils peuvent être facilement modifiés pour utiliser la `Worksheet.comments` propriété.</span><span class="sxs-lookup"><span data-stu-id="291df-116">The samples in this article work with comments at the workbook level, but they can be easily modified to use the `Worksheet.comments` property.</span></span>

## <a name="add-comments"></a><span data-ttu-id="291df-117">Ajouter des commentaires</span><span class="sxs-lookup"><span data-stu-id="291df-117">Add comments</span></span>

<span data-ttu-id="291df-118">Utilisez la `CommentCollection.add` méthode pour ajouter des commentaires à un classeur.</span><span class="sxs-lookup"><span data-stu-id="291df-118">Use the `CommentCollection.add` method to add comments to a workbook.</span></span> <span data-ttu-id="291df-119">Cette méthode peut prendre jusqu’à trois paramètres :</span><span class="sxs-lookup"><span data-stu-id="291df-119">This method takes up to three parameters:</span></span>

- <span data-ttu-id="291df-120">`cellAddress`: La cellule dans laquelle le commentaire est ajouté.</span><span class="sxs-lookup"><span data-stu-id="291df-120">`cellAddress`: The cell where the comment is added.</span></span> <span data-ttu-id="291df-121">Il peut s’agir d’un objet String ou [Range](/javascript/api/excel/excel.range) .</span><span class="sxs-lookup"><span data-stu-id="291df-121">This can either be a string or [Range](/javascript/api/excel/excel.range) object.</span></span> <span data-ttu-id="291df-122">La plage doit être une seule cellule.</span><span class="sxs-lookup"><span data-stu-id="291df-122">The range must be a single cell.</span></span>
- <span data-ttu-id="291df-123">`content`: Contenu du commentaire.</span><span class="sxs-lookup"><span data-stu-id="291df-123">`content`: The comment's content.</span></span> <span data-ttu-id="291df-124">Utilisez une chaîne pour les commentaires en texte brut.</span><span class="sxs-lookup"><span data-stu-id="291df-124">Use a string for plain text comments.</span></span> <span data-ttu-id="291df-125">Utilisez un objet [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) pour les commentaires avec des [mentions](#mentions).</span><span class="sxs-lookup"><span data-stu-id="291df-125">Use a [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) object for comments with [mentions](#mentions).</span></span>
- <span data-ttu-id="291df-126">`contentType`: Énumération [ContentType](/javascript/api/excel/excel.contenttype) spécifiant le type de contenu.</span><span class="sxs-lookup"><span data-stu-id="291df-126">`contentType`: A [ContentType](/javascript/api/excel/excel.contenttype) enum specifying type of content.</span></span> <span data-ttu-id="291df-127">La valeur par défaut est `ContentType.plain`.</span><span class="sxs-lookup"><span data-stu-id="291df-127">The default value is `ContentType.plain`.</span></span>

<span data-ttu-id="291df-128">L’exemple de code suivant ajoute un commentaire à la cellule **A2**.</span><span class="sxs-lookup"><span data-stu-id="291df-128">The following code sample adds a comment to cell **A2**.</span></span>

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
> <span data-ttu-id="291df-129">Les commentaires ajoutés par un complément sont attribués à l’utilisateur actuel de ce complément.</span><span class="sxs-lookup"><span data-stu-id="291df-129">Comments added by an add-in are attributed to the current user of that add-in.</span></span>

### <a name="add-comment-replies"></a><span data-ttu-id="291df-130">Ajouter des réponses aux commentaires</span><span class="sxs-lookup"><span data-stu-id="291df-130">Add comment replies</span></span>

<span data-ttu-id="291df-131">Un `Comment` objet est un thème de commentaire qui contient zéro ou plusieurs réponses.</span><span class="sxs-lookup"><span data-stu-id="291df-131">A `Comment` object is a comment thread that contains zero or more replies.</span></span> <span data-ttu-id="291df-132">Les objets `Comment` ont une propriété `replies`, qui est une collection [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) contenant des objets [CommentReply](/javascript/api/excel/excel.commentreply).</span><span class="sxs-lookup"><span data-stu-id="291df-132">`Comment` objects have a `replies` property, which is a [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) that contains [CommentReply](/javascript/api/excel/excel.commentreply) objects.</span></span> <span data-ttu-id="291df-133">Pour ajouter une réponse à un commentaire, utilisez la méthode `CommentReplyCollection.add`, en l’appliquant au texte de la réponse.</span><span class="sxs-lookup"><span data-stu-id="291df-133">To add a reply to a comment, use the `CommentReplyCollection.add` method, passing in the text of the reply.</span></span> <span data-ttu-id="291df-134">Les réponses s’affichent dans l’ordre dans lequel elles sont ajoutées.</span><span class="sxs-lookup"><span data-stu-id="291df-134">Replies are displayed in the order they are added.</span></span> <span data-ttu-id="291df-135">Elles sont également attribuées à l’utilisateur actuel du complément.</span><span class="sxs-lookup"><span data-stu-id="291df-135">They are also attributed to the current user of the add-in.</span></span>

<span data-ttu-id="291df-136">L’exemple de code suivant ajoute une réponse au premier commentaire du classeur.</span><span class="sxs-lookup"><span data-stu-id="291df-136">The following code sample adds a reply to the first comment in the workbook.</span></span>

```js
Excel.run(function (context) {
    // Get the first comment added to the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    return context.sync();
});
```

## <a name="edit-comments"></a><span data-ttu-id="291df-137">Modifier les commentaires</span><span class="sxs-lookup"><span data-stu-id="291df-137">Edit comments</span></span>

<span data-ttu-id="291df-138">Pour modifier un commentaire ou une réponse à un commentaire, configurez sa propriété `Comment.content` ou `CommentReply.content`.</span><span class="sxs-lookup"><span data-stu-id="291df-138">To edit a comment or comment reply, set its `Comment.content` property or `CommentReply.content` property.</span></span>

```js
Excel.run(function (context) {
    // Edit the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.content = "PLEASE add headers here.";
    return context.sync();
});
```

### <a name="edit-comment-replies"></a><span data-ttu-id="291df-139">Modifier les réponses de commentaire</span><span class="sxs-lookup"><span data-stu-id="291df-139">Edit comment replies</span></span>

<span data-ttu-id="291df-140">Pour modifier une réponse de commentaire, définissez sa `CommentReply.content` propriété.</span><span class="sxs-lookup"><span data-stu-id="291df-140">To edit a comment reply, set its `CommentReply.content` property.</span></span>

```js
Excel.run(function (context) {
    // Edit the first comment reply on the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    var reply = comment.replies.getItemAt(0);
    reply.content = "Never mind";
    return context.sync();
});
```

## <a name="delete-comments"></a><span data-ttu-id="291df-141">Supprimer les commentaires</span><span class="sxs-lookup"><span data-stu-id="291df-141">Delete comments</span></span>

<span data-ttu-id="291df-142">Pour supprimer un commentaire, utilisez la `Comment.delete` méthode.</span><span class="sxs-lookup"><span data-stu-id="291df-142">To delete a comment use the `Comment.delete` method.</span></span> <span data-ttu-id="291df-143">La suppression d’un commentaire supprime également les réponses associées à ce commentaire.</span><span class="sxs-lookup"><span data-stu-id="291df-143">Deleting a comment also deletes the replies associated with that comment.</span></span>

```js
Excel.run(function (context) {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    return context.sync();
});
```

### <a name="delete-comment-replies"></a><span data-ttu-id="291df-144">Supprimer les réponses de commentaire</span><span class="sxs-lookup"><span data-stu-id="291df-144">Delete comment replies</span></span>

<span data-ttu-id="291df-145">Pour supprimer une réponse de commentaire, utilisez la `CommentReply.delete` méthode.</span><span class="sxs-lookup"><span data-stu-id="291df-145">To delete a comment reply, use the `CommentReply.delete` method.</span></span>

```js
Excel.run(function (context) {
    // Delete the first comment reply from this worksheet's first comment.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="resolve-comment-threads"></a><span data-ttu-id="291df-146">Résoudre les thèmes de commentaires</span><span class="sxs-lookup"><span data-stu-id="291df-146">Resolve comment threads</span></span>

<span data-ttu-id="291df-147">Un thread de commentaire a une valeur booléenne configurable, `resolved` pour indiquer s’il est résolu.</span><span class="sxs-lookup"><span data-stu-id="291df-147">A comment thread has a configurable boolean value, `resolved`, to indicate if it is resolved.</span></span> <span data-ttu-id="291df-148">Une valeur de `true` signifie que le thread de commentaire est résolu.</span><span class="sxs-lookup"><span data-stu-id="291df-148">A value of `true` means the comment thread is resolved.</span></span> <span data-ttu-id="291df-149">Une valeur de `false` signifie que le fil de commentaires est nouveau ou rouvert.</span><span class="sxs-lookup"><span data-stu-id="291df-149">A value of `false` means the comment thread is either new or reopened.</span></span>

```js
Excel.run(function (context) {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    return context.sync();
});
```

<span data-ttu-id="291df-150">Les réponses de commentaire ont une `resolved` propriété ReadOnly.</span><span class="sxs-lookup"><span data-stu-id="291df-150">Comment replies have a readonly `resolved` property.</span></span> <span data-ttu-id="291df-151">Sa valeur est toujours égale à celle du reste du thread.</span><span class="sxs-lookup"><span data-stu-id="291df-151">Its value is always equal to that of the rest of the thread.</span></span>

## <a name="comment-metadata"></a><span data-ttu-id="291df-152">Métadonnées de commentaire</span><span class="sxs-lookup"><span data-stu-id="291df-152">Comment metadata</span></span>

<span data-ttu-id="291df-153">Chaque commentaire contient des métadonnées concernant sa création, notamment l’auteur et la date de création.</span><span class="sxs-lookup"><span data-stu-id="291df-153">Each comment contains metadata about its creation, such as the author and creation date.</span></span> <span data-ttu-id="291df-154">Les commentaires créés par votre complément sont considérés comme créés par l’utilisateur actuel.</span><span class="sxs-lookup"><span data-stu-id="291df-154">Comments created by your add-in are considered to be authored by the current user.</span></span>

<span data-ttu-id="291df-155">L’exemple suivant montre comment afficher l’adresse e-mail et le nom de l’auteur, ainsi que la date de création d’un commentaire dans la cellule **A2**.</span><span class="sxs-lookup"><span data-stu-id="291df-155">The following sample shows how to display the author's email, author's name, and creation date of a comment at **A2**.</span></span>

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

### <a name="comment-reply-metadata"></a><span data-ttu-id="291df-156">Métadonnées de réponse de commentaire</span><span class="sxs-lookup"><span data-stu-id="291df-156">Comment reply metadata</span></span>

<span data-ttu-id="291df-157">Les réponses aux commentaires stockent les mêmes types de métadonnées que le commentaire initial.</span><span class="sxs-lookup"><span data-stu-id="291df-157">Comment replies store the same types of metadata as the initial comment.</span></span>

<span data-ttu-id="291df-158">L’exemple suivant montre comment afficher le courrier électronique, le nom de l’auteur et la date de création de l’auteur de la réponse de commentaire la plus récente à la version **a2**.</span><span class="sxs-lookup"><span data-stu-id="291df-158">The following sample shows how to display the author's email, author's name, and creation date of the latest comment reply at **A2**.</span></span>

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

## <a name="mentions"></a><span data-ttu-id="291df-159">Mentions</span><span class="sxs-lookup"><span data-stu-id="291df-159">Mentions</span></span>

<span data-ttu-id="291df-160">Les [mentions](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd) sont utilisées pour marquer les collègues dans un commentaire.</span><span class="sxs-lookup"><span data-stu-id="291df-160">[Mentions](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd) are used to tag colleagues in a comment.</span></span> <span data-ttu-id="291df-161">Les notifications sont envoyées avec le contenu de votre commentaire.</span><span class="sxs-lookup"><span data-stu-id="291df-161">This sends them notifications with your comment's content.</span></span> <span data-ttu-id="291df-162">Votre complément peut créer ces mentions à votre place.</span><span class="sxs-lookup"><span data-stu-id="291df-162">Your add-in can create these mentions on your behalf.</span></span>

<span data-ttu-id="291df-163">Les commentaires avec des mentions doivent être créés avec des objets [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) .</span><span class="sxs-lookup"><span data-stu-id="291df-163">Comments with mentions need to be created with [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) objects.</span></span> <span data-ttu-id="291df-164">Appelez `CommentCollection.add` avec un `CommentRichContent` conteneur contenant une ou plusieurs mentions et spécifiez `ContentType.mention` comme `contentType` paramètre.</span><span class="sxs-lookup"><span data-stu-id="291df-164">Call `CommentCollection.add` with a `CommentRichContent` containing one or more mentions and specify `ContentType.mention` as the `contentType` parameter.</span></span> <span data-ttu-id="291df-165">La `content` chaîne doit également être mise en forme pour insérer la mention dans le texte.</span><span class="sxs-lookup"><span data-stu-id="291df-165">The `content` string also needs to be formatted to insert the mention into the text.</span></span> <span data-ttu-id="291df-166">Le format d’une mention est le suivant : `<at id="{replyIndex}">{mentionName}</at>` .</span><span class="sxs-lookup"><span data-stu-id="291df-166">The format for a mention is: `<at id="{replyIndex}">{mentionName}</at>`.</span></span>

> [!NOTE]
> <span data-ttu-id="291df-167">Actuellement, seul le nom exact de la mention peut être utilisé comme texte du lien mention.</span><span class="sxs-lookup"><span data-stu-id="291df-167">Currently, only the mention's exact name can be used as the text of the mention link.</span></span> <span data-ttu-id="291df-168">La prise en charge des versions raccourcies d’un nom sera ajoutée ultérieurement.</span><span class="sxs-lookup"><span data-stu-id="291df-168">Support for shortened versions of a name will be added later.</span></span>

<span data-ttu-id="291df-169">L’exemple suivant montre un commentaire avec une seule mention.</span><span class="sxs-lookup"><span data-stu-id="291df-169">The following example shows a comment with a single mention.</span></span>

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

## <a name="comment-events"></a><span data-ttu-id="291df-170">Événements de commentaire</span><span class="sxs-lookup"><span data-stu-id="291df-170">Comment events</span></span>

<span data-ttu-id="291df-171">Votre complément peut écouter les ajouts, les modifications et les suppressions de commentaires.</span><span class="sxs-lookup"><span data-stu-id="291df-171">Your add-in can listen for comment additions, changes, and deletions.</span></span> <span data-ttu-id="291df-172">Les [événements de commentaire](/javascript/api/excel/excel.commentcollection#event-details) se produisent sur l' `CommentCollection` objet.</span><span class="sxs-lookup"><span data-stu-id="291df-172">[Comment events](/javascript/api/excel/excel.commentcollection#event-details) occur on the `CommentCollection` object.</span></span> <span data-ttu-id="291df-173">Pour écouter les événements de commentaire, enregistrez `onAdded` le `onChanged` Gestionnaire d’événements,, ou le `onDeleted` commentaire.</span><span class="sxs-lookup"><span data-stu-id="291df-173">To listen for comment events, register the `onAdded`, `onChanged`, or `onDeleted` comment event handler.</span></span> <span data-ttu-id="291df-174">Lorsqu’un événement de commentaire est détecté, utilisez ce gestionnaire d’événements pour récupérer des données sur le Commentaire ajouté, modifié ou supprimé.</span><span class="sxs-lookup"><span data-stu-id="291df-174">When a comment event is detected, use this event handler to retrieve data about the added, changed, or deleted comment.</span></span> <span data-ttu-id="291df-175">L' `onChanged` événement gère également les ajouts de réponse aux commentaires, les modifications et les suppressions.</span><span class="sxs-lookup"><span data-stu-id="291df-175">The `onChanged` event also handles comment reply additions, changes, and deletions.</span></span> 

<span data-ttu-id="291df-176">Chaque événement de commentaire ne déclenche qu’une seule fois lorsque plusieurs ajouts, modifications ou suppressions sont effectués en même temps.</span><span class="sxs-lookup"><span data-stu-id="291df-176">Each comment event only triggers once when multiple additions, changes, or deletions are performed at the same time.</span></span> <span data-ttu-id="291df-177">Tous les objets [CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs), [CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventarg)et [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) contiennent des tableaux d’ID de commentaires permettant de mapper les actions d’événement vers les collections de commentaires.</span><span class="sxs-lookup"><span data-stu-id="291df-177">All the [CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs), [CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventarg), and [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) objects contain arrays of comment IDs to map the event actions back to the comment collections.</span></span>

<span data-ttu-id="291df-178">Pour plus d’informations sur l’inscription de gestionnaires d’événements, la gestion des événements et la suppression de gestionnaires d’événements, voir l’article [work with Events using the Excel JavaScript API](excel-add-ins-events.md) .</span><span class="sxs-lookup"><span data-stu-id="291df-178">See the [Work with Events using the Excel JavaScript API](excel-add-ins-events.md) article for additional information about registering event handlers, handling events, and removing event handlers.</span></span> 

### <a name="comment-addition-events"></a><span data-ttu-id="291df-179">Événements d’ajout de commentaires</span><span class="sxs-lookup"><span data-stu-id="291df-179">Comment addition events</span></span> 
<span data-ttu-id="291df-180">L' `onAdded` événement est déclenché lorsqu’un ou plusieurs nouveaux commentaires sont ajoutés à la collection de commentaires.</span><span class="sxs-lookup"><span data-stu-id="291df-180">The `onAdded` event is triggered when one or more new comments are added to the comment collection.</span></span> <span data-ttu-id="291df-181">Cet événement n’est *pas* déclenché lorsque les réponses sont ajoutées à un thread de commentaire (voir [événements de modification](#comment-change-events) des commentaires pour en savoir plus sur les événements de réponse aux commentaires).</span><span class="sxs-lookup"><span data-stu-id="291df-181">This event is *not* triggered when replies are added to a comment thread (see [Comment change events](#comment-change-events) to learn about comment reply events).</span></span>

<span data-ttu-id="291df-182">L’exemple suivant montre comment inscrire le `onAdded` Gestionnaire d’événements, puis utiliser l' `CommentAddedEventArgs` objet pour récupérer le `commentDetails` tableau du Commentaire ajouté.</span><span class="sxs-lookup"><span data-stu-id="291df-182">The following sample shows how to register the `onAdded` event handler and then use the `CommentAddedEventArgs` object to retrieve the `commentDetails` array of the added comment.</span></span>

> [!NOTE]
> <span data-ttu-id="291df-183">Cet exemple fonctionne uniquement lorsqu’un seul commentaire est ajouté.</span><span class="sxs-lookup"><span data-stu-id="291df-183">This sample only works when a single comment is added.</span></span> 

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

### <a name="comment-change-events"></a><span data-ttu-id="291df-184">Événements de modification de commentaire</span><span class="sxs-lookup"><span data-stu-id="291df-184">Comment change events</span></span> 
<span data-ttu-id="291df-185">L' `onChanged` événement comment est déclenché dans les scénarios suivants.</span><span class="sxs-lookup"><span data-stu-id="291df-185">The `onChanged` comment event is triggered in the following scenarios.</span></span>

- <span data-ttu-id="291df-186">Le contenu d’un commentaire est mis à jour.</span><span class="sxs-lookup"><span data-stu-id="291df-186">A comment's content is updated.</span></span>
- <span data-ttu-id="291df-187">Une thread de commentaire est résolue.</span><span class="sxs-lookup"><span data-stu-id="291df-187">A comment thread is resolved.</span></span>
- <span data-ttu-id="291df-188">Une thread de commentaire est rouverte.</span><span class="sxs-lookup"><span data-stu-id="291df-188">A comment thread is reopened.</span></span>
- <span data-ttu-id="291df-189">Une réponse est ajoutée à une thread de commentaire.</span><span class="sxs-lookup"><span data-stu-id="291df-189">A reply is added to a comment thread.</span></span>
- <span data-ttu-id="291df-190">Une réponse est mise à jour dans une thread de commentaire.</span><span class="sxs-lookup"><span data-stu-id="291df-190">A reply is updated in a comment thread.</span></span>
- <span data-ttu-id="291df-191">Une réponse est supprimée dans une thread de commentaire.</span><span class="sxs-lookup"><span data-stu-id="291df-191">A reply is deleted in a comment thread.</span></span>

<span data-ttu-id="291df-192">L’exemple suivant montre comment inscrire le `onChanged` Gestionnaire d’événements, puis utiliser l' `CommentChangedEventArgs` objet pour récupérer le `commentDetails` tableau du commentaire modifié.</span><span class="sxs-lookup"><span data-stu-id="291df-192">The following sample shows how to register the `onChanged` event handler and then use the `CommentChangedEventArgs` object to retrieve the `commentDetails` array of the changed comment.</span></span>

> [!NOTE]
> <span data-ttu-id="291df-193">Cet exemple fonctionne uniquement lorsqu’un seul commentaire est modifié.</span><span class="sxs-lookup"><span data-stu-id="291df-193">This sample only works when a single comment is changed.</span></span> 

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

### <a name="comment-deletion-events"></a><span data-ttu-id="291df-194">Événements de suppression de commentaires</span><span class="sxs-lookup"><span data-stu-id="291df-194">Comment deletion events</span></span>
<span data-ttu-id="291df-195">L' `onDeleted` événement est déclenché lorsqu’un commentaire est supprimé de la collection de commentaires.</span><span class="sxs-lookup"><span data-stu-id="291df-195">The `onDeleted` event is triggered when a comment is deleted from the comment collection.</span></span> <span data-ttu-id="291df-196">Une fois qu’un commentaire a été supprimé, ses métadonnées ne sont plus disponibles.</span><span class="sxs-lookup"><span data-stu-id="291df-196">Once a comment has been deleted, its metadata is no longer available.</span></span> <span data-ttu-id="291df-197">L’objet [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) fournit des ID de commentaire, si votre complément gère des Commentaires individuels.</span><span class="sxs-lookup"><span data-stu-id="291df-197">The [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) object provides comment IDs, in case your add-in is managing individual comments.</span></span>

<span data-ttu-id="291df-198">L’exemple suivant montre comment inscrire le `onDeleted` Gestionnaire d’événements, puis utiliser l' `CommentDeletedEventArgs` objet pour récupérer le `commentDetails` tableau du commentaire supprimé.</span><span class="sxs-lookup"><span data-stu-id="291df-198">The following sample shows how to register the `onDeleted` event handler and then use the `CommentDeletedEventArgs` object to retrieve the `commentDetails` array of the deleted comment.</span></span>

> [!NOTE]
> <span data-ttu-id="291df-199">Cet exemple ne fonctionne qu’en cas de suppression d’un seul commentaire.</span><span class="sxs-lookup"><span data-stu-id="291df-199">This sample only works when a single comment is deleted.</span></span> 

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

## <a name="see-also"></a><span data-ttu-id="291df-200">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="291df-200">See also</span></span>

- [<span data-ttu-id="291df-201">Modèle objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="291df-201">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="291df-202">Utiliser les classeurs utilisant l’API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="291df-202">Work with workbooks using the Excel JavaScript API</span></span>](excel-add-ins-workbooks.md)
- [<span data-ttu-id="291df-203">Utilisation d’événements à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="291df-203">Work with Events using the Excel JavaScript API</span></span>](excel-add-ins-events.md)
- [<span data-ttu-id="291df-204">Insérer des commentaires et des notes dans Excel</span><span class="sxs-lookup"><span data-stu-id="291df-204">Insert comments and notes in Excel</span></span>](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
