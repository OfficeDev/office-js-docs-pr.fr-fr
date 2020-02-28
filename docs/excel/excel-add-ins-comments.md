---
title: Utiliser des commentaires à l’aide de l’API JavaScript pour Excel
description: Informations sur l’utilisation des API pour ajouter, supprimer et modifier des commentaires et des thèmes de commentaires.
ms.date: 02/11/2020
localization_priority: Normal
ms.openlocfilehash: d6be0f07e0d3bb134385f0a08c20ce00da4de892
ms.sourcegitcommit: d85efbf41a3382ca7d3ab08f2c3f0664d4b26c53
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/28/2020
ms.locfileid: "42327781"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a><span data-ttu-id="bc736-103">Utiliser des commentaires à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="bc736-103">Work with comments using the Excel JavaScript API</span></span>

<span data-ttu-id="bc736-104">Cet article explique comment ajouter, lire, modifier et supprimer des commentaires dans un classeur à l’aide de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="bc736-104">This article describes how to add, read, modify, and remove comments in a workbook with the Excel JavaScript API.</span></span> <span data-ttu-id="bc736-105">Pour en savoir plus sur la fonctionnalité de commentaire, consultez l’article [Insérer des commentaires et des notes dans Excel](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) .</span><span class="sxs-lookup"><span data-stu-id="bc736-105">You can learn more about the comment feature from the [Insert comments and notes in Excel](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) article.</span></span>

<span data-ttu-id="bc736-106">Dans l’API JavaScript pour Excel, un commentaire est à la fois la note initiale et la discussion thématique connectée.</span><span class="sxs-lookup"><span data-stu-id="bc736-106">In the Excel JavaScript API, a comment is both the initial note and the connected threaded discussion.</span></span> <span data-ttu-id="bc736-107">Elle est liée à une cellule individuelle.</span><span class="sxs-lookup"><span data-stu-id="bc736-107">It is tied to an individual cell.</span></span> <span data-ttu-id="bc736-108">Toute personne qui consulte le classeur avec des autorisations suffisantes peut répondre à un commentaire.</span><span class="sxs-lookup"><span data-stu-id="bc736-108">Anyone viewing the workbook with sufficient permissions can reply to a comment.</span></span> <span data-ttu-id="bc736-109">Un objet [Comment](/javascript/api/excel/excel.comment) stocke ces réponses en tant qu’objets [CommentReply](/javascript/api/excel/excel.commentreply) .</span><span class="sxs-lookup"><span data-stu-id="bc736-109">A [Comment](/javascript/api/excel/excel.comment) object stores those replies as [CommentReply](/javascript/api/excel/excel.commentreply) objects.</span></span> <span data-ttu-id="bc736-110">Vous devez considérer un commentaire comme un fil de discussion et qu’un thread doit avoir une entrée spéciale comme point de départ.</span><span class="sxs-lookup"><span data-stu-id="bc736-110">You should consider a comment to be a thread and that a thread must have a special entry as the starting point.</span></span>

![Commentaire Excel, étiqueté « commentaire » avec deux réponses, intitulées « comment. réponses [0] » et «comment. réponses [1].](../images/excel-comments.png)

<span data-ttu-id="bc736-112">Les commentaires d’un classeur sont suivis `Workbook.comments` par la propriété.</span><span class="sxs-lookup"><span data-stu-id="bc736-112">Comments within a workbook are tracked by the `Workbook.comments` property.</span></span> <span data-ttu-id="bc736-113">Cela inclut les commentaires créés par les utilisateurs ainsi que les commentaires créés par votre complément.</span><span class="sxs-lookup"><span data-stu-id="bc736-113">This includes comments created by users and also comments created by your add-in.</span></span> <span data-ttu-id="bc736-114">La propriété `Workbook.comments` est un objet [CommentCollection](/javascript/api/excel/excel.commentcollection) qui contient une collection d’objets [Comment](/javascript/api/excel/excel.comment).</span><span class="sxs-lookup"><span data-stu-id="bc736-114">The `Workbook.comments` property is a [CommentCollection](/javascript/api/excel/excel.commentcollection) object that contains a collection of [Comment](/javascript/api/excel/excel.comment) objects.</span></span> <span data-ttu-id="bc736-115">Les commentaires sont également accessibles au niveau de la [feuille de calcul](/javascript/api/excel/excel.worksheet) .</span><span class="sxs-lookup"><span data-stu-id="bc736-115">Comments are also accessible at the [Worksheet](/javascript/api/excel/excel.worksheet) level.</span></span> <span data-ttu-id="bc736-116">Les exemples de cet article utilisent des commentaires au niveau du classeur, mais ils peuvent être facilement modifiés pour utiliser `Worksheet.comments` la propriété.</span><span class="sxs-lookup"><span data-stu-id="bc736-116">The samples in this article work with comments at the workbook level, but they can be easily modified to use the `Worksheet.comments` property.</span></span>

## <a name="add-comments"></a><span data-ttu-id="bc736-117">Ajouter des commentaires</span><span class="sxs-lookup"><span data-stu-id="bc736-117">Add comments</span></span>

<span data-ttu-id="bc736-118">Utilisez la `CommentCollection.add` méthode pour ajouter des commentaires à un classeur.</span><span class="sxs-lookup"><span data-stu-id="bc736-118">Use the `CommentCollection.add` method to add comments to a workbook.</span></span> <span data-ttu-id="bc736-119">Cette méthode peut prendre jusqu’à trois paramètres :</span><span class="sxs-lookup"><span data-stu-id="bc736-119">This method takes up to three parameters:</span></span>

- <span data-ttu-id="bc736-120">`cellAddress`: La cellule dans laquelle le commentaire est ajouté.</span><span class="sxs-lookup"><span data-stu-id="bc736-120">`cellAddress`: The cell where the comment is added.</span></span> <span data-ttu-id="bc736-121">Il peut s’agir d’un objet String ou [Range](/javascript/api/excel/excel.range) .</span><span class="sxs-lookup"><span data-stu-id="bc736-121">This can either be a string or [Range](/javascript/api/excel/excel.range) object.</span></span> <span data-ttu-id="bc736-122">La plage doit être une seule cellule.</span><span class="sxs-lookup"><span data-stu-id="bc736-122">The range must be a single cell.</span></span>
- <span data-ttu-id="bc736-123">`content`: Contenu du commentaire.</span><span class="sxs-lookup"><span data-stu-id="bc736-123">`content`: The comment's content.</span></span> <span data-ttu-id="bc736-124">Utilisez une chaîne pour les commentaires en texte brut.</span><span class="sxs-lookup"><span data-stu-id="bc736-124">Use a string for plain text comments.</span></span> <span data-ttu-id="bc736-125">Utilisez un objet [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) pour les commentaires avec des [mentions](#mentions-online-only).</span><span class="sxs-lookup"><span data-stu-id="bc736-125">Use a [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) object for comments with [mentions](#mentions-online-only).</span></span> 
- <span data-ttu-id="bc736-126">`contentType`: Énumération [ContentType](/javascript/api/excel/excel.contenttype) spécifiant le type de contenu.</span><span class="sxs-lookup"><span data-stu-id="bc736-126">`contentType`: A [ContentType](/javascript/api/excel/excel.contenttype) enum specifying type of content.</span></span> <span data-ttu-id="bc736-127">La valeur par défaut est `ContentType.plain`.</span><span class="sxs-lookup"><span data-stu-id="bc736-127">The default value is `ContentType.plain`.</span></span>

<span data-ttu-id="bc736-128">L’exemple de code suivant ajoute un commentaire à la cellule **A2**.</span><span class="sxs-lookup"><span data-stu-id="bc736-128">The following code sample adds a comment to cell **A2**.</span></span>

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
> <span data-ttu-id="bc736-129">Les commentaires ajoutés par un complément sont attribués à l’utilisateur actuel de ce complément.</span><span class="sxs-lookup"><span data-stu-id="bc736-129">Comments added by an add-in are attributed to the current user of that add-in.</span></span>

### <a name="add-comment-replies"></a><span data-ttu-id="bc736-130">Ajouter des réponses aux commentaires</span><span class="sxs-lookup"><span data-stu-id="bc736-130">Add comment replies</span></span>

<span data-ttu-id="bc736-131">Un `Comment` objet est un thème de commentaire qui contient zéro ou plusieurs réponses.</span><span class="sxs-lookup"><span data-stu-id="bc736-131">A `Comment` object is a comment thread that contains zero or more replies.</span></span> <span data-ttu-id="bc736-132">Les objets `Comment` ont une propriété `replies`, qui est une collection [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) contenant des objets [CommentReply](/javascript/api/excel/excel.commentreply).</span><span class="sxs-lookup"><span data-stu-id="bc736-132">`Comment` objects have a `replies` property, which is a [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) that contains [CommentReply](/javascript/api/excel/excel.commentreply) objects.</span></span> <span data-ttu-id="bc736-133">Pour ajouter une réponse à un commentaire, utilisez la méthode `CommentReplyCollection.add`, en l’appliquant au texte de la réponse.</span><span class="sxs-lookup"><span data-stu-id="bc736-133">To add a reply to a comment, use the `CommentReplyCollection.add` method, passing in the text of the reply.</span></span> <span data-ttu-id="bc736-134">Les réponses s’affichent dans l’ordre dans lequel elles sont ajoutées.</span><span class="sxs-lookup"><span data-stu-id="bc736-134">Replies are displayed in the order they are added.</span></span> <span data-ttu-id="bc736-135">Elles sont également attribuées à l’utilisateur actuel du complément.</span><span class="sxs-lookup"><span data-stu-id="bc736-135">They are also attributed to the current user of the add-in.</span></span>

<span data-ttu-id="bc736-136">L’exemple de code suivant ajoute une réponse au premier commentaire du classeur.</span><span class="sxs-lookup"><span data-stu-id="bc736-136">The following code sample adds a reply to the first comment in the workbook.</span></span>

```js
Excel.run(function (context) {
    // Get the first comment added to the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    return context.sync();
});
```

## <a name="edit-comments"></a><span data-ttu-id="bc736-137">Modifier les commentaires</span><span class="sxs-lookup"><span data-stu-id="bc736-137">Edit comments</span></span>

<span data-ttu-id="bc736-138">Pour modifier un commentaire ou une réponse à un commentaire, configurez sa propriété `Comment.content` ou `CommentReply.content`.</span><span class="sxs-lookup"><span data-stu-id="bc736-138">To edit a comment or comment reply, set its `Comment.content` property or `CommentReply.content` property.</span></span>

```js
Excel.run(function (context) {
    // Edit the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.content = "PLEASE add headers here.";
    return context.sync();
});
```

### <a name="edit-comment-replies"></a><span data-ttu-id="bc736-139">Modifier les réponses de commentaire</span><span class="sxs-lookup"><span data-stu-id="bc736-139">Edit comment replies</span></span>

<span data-ttu-id="bc736-140">Pour modifier une réponse de commentaire, définissez `CommentReply.content` sa propriété.</span><span class="sxs-lookup"><span data-stu-id="bc736-140">To edit a comment reply, set its `CommentReply.content` property.</span></span>

```js
Excel.run(function (context) {
    // Edit the first comment reply on the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    var reply = comment.replies.getItemAt(0);
    reply.content = "Never mind";
    return context.sync();
});
```

## <a name="delete-comments"></a><span data-ttu-id="bc736-141">Supprimer les commentaires</span><span class="sxs-lookup"><span data-stu-id="bc736-141">Delete comments</span></span>

<span data-ttu-id="bc736-142">Pour supprimer un commentaire, utilisez `Comment.delete` la méthode.</span><span class="sxs-lookup"><span data-stu-id="bc736-142">To delete a comment use the `Comment.delete` method.</span></span> <span data-ttu-id="bc736-143">La suppression d’un commentaire supprime également les réponses associées à ce commentaire.</span><span class="sxs-lookup"><span data-stu-id="bc736-143">Deleting a comment also deletes the replies associated with that comment.</span></span>

```js
Excel.run(function (context) {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    return context.sync();
});
```

### <a name="delete-comment-replies"></a><span data-ttu-id="bc736-144">Supprimer les réponses de commentaire</span><span class="sxs-lookup"><span data-stu-id="bc736-144">Delete comment replies</span></span>

<span data-ttu-id="bc736-145">Pour supprimer une réponse de commentaire, utilisez `CommentReply.delete` la méthode.</span><span class="sxs-lookup"><span data-stu-id="bc736-145">To delete a comment reply, use the `CommentReply.delete` method.</span></span>

```js
Excel.run(function (context) {
    // Delete the first comment reply from this worksheet's first comment.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="resolve-comment-threads-preview"></a><span data-ttu-id="bc736-146">Résoudre les thèmes de commentaire ([Aperçu](../reference/requirement-sets/excel-preview-apis.md))</span><span class="sxs-lookup"><span data-stu-id="bc736-146">Resolve comment threads ([preview](../reference/requirement-sets/excel-preview-apis.md))</span></span> 

<span data-ttu-id="bc736-147">Un thread de commentaire a une valeur `resolved`booléenne configurable, pour indiquer s’il est résolu.</span><span class="sxs-lookup"><span data-stu-id="bc736-147">A comment thread has a configurable boolean value, `resolved`, to indicate if it is resolved.</span></span> <span data-ttu-id="bc736-148">Une valeur de `true` signifie que le thread de commentaire est résolu.</span><span class="sxs-lookup"><span data-stu-id="bc736-148">A value of `true` means the comment thread is resolved.</span></span> <span data-ttu-id="bc736-149">Une valeur de `false` signifie que le fil de commentaires est nouveau ou rouvert.</span><span class="sxs-lookup"><span data-stu-id="bc736-149">A value of `false` means the comment thread is either new or reopened.</span></span>

```js
Excel.run(function (context) {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    return context.sync();
});
```

<span data-ttu-id="bc736-150">Les réponses de commentaire ont `resolved` une propriété ReadOnly.</span><span class="sxs-lookup"><span data-stu-id="bc736-150">Comment replies have a readonly `resolved` property.</span></span> <span data-ttu-id="bc736-151">Sa valeur est toujours égale à celle du reste du thread.</span><span class="sxs-lookup"><span data-stu-id="bc736-151">Its value is always equal to that of the rest of the thread.</span></span>

## <a name="comment-metadata"></a><span data-ttu-id="bc736-152">Métadonnées de commentaire</span><span class="sxs-lookup"><span data-stu-id="bc736-152">Comment metadata</span></span>

<span data-ttu-id="bc736-153">Chaque commentaire contient des métadonnées concernant sa création, notamment l’auteur et la date de création.</span><span class="sxs-lookup"><span data-stu-id="bc736-153">Each comment contains metadata about its creation, such as the author and creation date.</span></span> <span data-ttu-id="bc736-154">Les commentaires créés par votre complément sont considérés comme créés par l’utilisateur actuel.</span><span class="sxs-lookup"><span data-stu-id="bc736-154">Comments created by your add-in are considered to be authored by the current user.</span></span>

<span data-ttu-id="bc736-155">L’exemple suivant montre comment afficher l’adresse e-mail et le nom de l’auteur, ainsi que la date de création d’un commentaire dans la cellule **A2**.</span><span class="sxs-lookup"><span data-stu-id="bc736-155">The following sample shows how to display the author's email, author's name, and creation date of a comment at **A2**.</span></span>

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

### <a name="comment-reply-metadata"></a><span data-ttu-id="bc736-156">Métadonnées de réponse de commentaire</span><span class="sxs-lookup"><span data-stu-id="bc736-156">Comment reply metadata</span></span>

<span data-ttu-id="bc736-157">Les réponses aux commentaires stockent les mêmes types de métadonnées que le commentaire initial.</span><span class="sxs-lookup"><span data-stu-id="bc736-157">Comment replies store the same types of metadata as the initial comment.</span></span>

<span data-ttu-id="bc736-158">L’exemple suivant montre comment afficher le courrier électronique, le nom de l’auteur et la date de création de l’auteur de la réponse de commentaire la plus récente à la version **a2**.</span><span class="sxs-lookup"><span data-stu-id="bc736-158">The following sample shows how to display the author's email, author's name, and creation date of the latest comment reply at **A2**.</span></span>

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

## <a name="mentions-online-only"></a><span data-ttu-id="bc736-159">Mentions ([en ligne uniquement](../reference/requirement-sets/excel-api-online-requirement-set.md))</span><span class="sxs-lookup"><span data-stu-id="bc736-159">Mentions ([online-only](../reference/requirement-sets/excel-api-online-requirement-set.md))</span></span> 

> [!NOTE]
> <span data-ttu-id="bc736-160">Le commentaire mentionne les API sont actuellement disponibles uniquement en préversion publique.</span><span class="sxs-lookup"><span data-stu-id="bc736-160">The comment mention APIs are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

> [!IMPORTANT]
> <span data-ttu-id="bc736-161">Les mentions de commentaire sont actuellement uniquement prises en charge pour Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="bc736-161">Comment mentions are currently only supported for Excel on the web.</span></span>

<span data-ttu-id="bc736-162">Les [mentions](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd) sont utilisées pour marquer les collègues dans un commentaire.</span><span class="sxs-lookup"><span data-stu-id="bc736-162">[Mentions](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd) are used to tag colleagues in a comment.</span></span> <span data-ttu-id="bc736-163">Les notifications sont envoyées avec le contenu de votre commentaire.</span><span class="sxs-lookup"><span data-stu-id="bc736-163">This sends them notifications with your comment's content.</span></span> <span data-ttu-id="bc736-164">Votre complément peut créer ces mentions à votre place.</span><span class="sxs-lookup"><span data-stu-id="bc736-164">Your add-in can create these mentions on your behalf.</span></span>

<span data-ttu-id="bc736-165">Les commentaires avec des mentions doivent être créés avec des objets [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) .</span><span class="sxs-lookup"><span data-stu-id="bc736-165">Comments with mentions need to be created with [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) objects.</span></span> <span data-ttu-id="bc736-166">Appelez `CommentCollection.add` avec un `CommentRichContent` conteneur contenant une ou plusieurs mentions et `ContentType.mention` spécifiez `contentType` comme paramètre.</span><span class="sxs-lookup"><span data-stu-id="bc736-166">Call `CommentCollection.add` with a `CommentRichContent` containing one or more mentions and specify `ContentType.mention` as the `contentType` parameter.</span></span> <span data-ttu-id="bc736-167">La `content` chaîne doit également être mise en forme pour insérer la mention dans le texte.</span><span class="sxs-lookup"><span data-stu-id="bc736-167">The `content` string also needs to be formatted to insert the mention into the text.</span></span> <span data-ttu-id="bc736-168">Le format d’une mention est le `<at id="{replyIndex}">{mentionName}</at>`suivant :.</span><span class="sxs-lookup"><span data-stu-id="bc736-168">The format for a mention is: `<at id="{replyIndex}">{mentionName}</at>`.</span></span>

> <span data-ttu-id="bc736-169">Note Actuellement, seul le nom exact de la mention peut être utilisé comme texte du lien mention.</span><span class="sxs-lookup"><span data-stu-id="bc736-169">[NOTE] Currently, only the mention's exact name can be used as the text of the mention link.</span></span> <span data-ttu-id="bc736-170">La prise en charge des versions raccourcies d’un nom sera ajoutée ultérieurement.</span><span class="sxs-lookup"><span data-stu-id="bc736-170">Support for shortened versions of a name will be added later.</span></span>

<span data-ttu-id="bc736-171">L’exemple suivant montre un commentaire avec une seule mention.</span><span class="sxs-lookup"><span data-stu-id="bc736-171">The following example shows a comment with a single mention.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="bc736-172">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="bc736-172">See also</span></span>

- [<span data-ttu-id="bc736-173">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="bc736-173">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="bc736-174">Utiliser les classeurs utilisant l’API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="bc736-174">Work with workbooks using the Excel JavaScript API</span></span>](excel-add-ins-workbooks.md)
- [<span data-ttu-id="bc736-175">Insérer des commentaires et des notes dans Excel</span><span class="sxs-lookup"><span data-stu-id="bc736-175">Insert comments and notes in Excel</span></span>](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
