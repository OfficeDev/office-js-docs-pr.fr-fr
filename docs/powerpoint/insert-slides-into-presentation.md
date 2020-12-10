---
title: Insertion et suppression de diapositives dans une présentation PowerPoint
description: Découvrez comment insérer des diapositives d’une présentation dans une autre et comment supprimer des diapositives.
ms.date: 12/04/2020
localization_priority: Normal
ms.openlocfilehash: ceb78054a95ac4b26bd71f79a086a00e3dce5278
ms.sourcegitcommit: cba180ae712d88d8d9ec417b4d1c7112cd8fdd17
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/09/2020
ms.locfileid: "49613703"
---
# <a name="insert-and-delete-slides-in-a-powerpoint-presentation-preview"></a><span data-ttu-id="b92e3-103">Insertion et suppression de diapositives dans une présentation PowerPoint (aperçu)</span><span class="sxs-lookup"><span data-stu-id="b92e3-103">Insert and delete slides in a PowerPoint presentation (preview)</span></span>

<span data-ttu-id="b92e3-104">Un complément PowerPoint peut insérer des diapositives d’une présentation dans la présentation en cours à l’aide de la bibliothèque JavaScript propre à l’application de PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="b92e3-104">A PowerPoint add-in can insert slides from one presentation into the current presentation by using PowerPoint's application-specific JavaScript library.</span></span> <span data-ttu-id="b92e3-105">Vous pouvez contrôler si les diapositives insérées conservent la mise en forme de la présentation source ou la mise en forme de la présentation cible.</span><span class="sxs-lookup"><span data-stu-id="b92e3-105">You can control whether the inserted slides keep the formatting of the source presentation or the formatting of the target presentation.</span></span> <span data-ttu-id="b92e3-106">Vous pouvez également supprimer des diapositives de la présentation.</span><span class="sxs-lookup"><span data-stu-id="b92e3-106">You can also delete slides from the presentation.</span></span>

[!include[General preview API prerequisites](../includes/using-preview-apis-host.md)]

<span data-ttu-id="b92e3-107">Les API d’insertion de diapositives sont principalement utilisées dans les scénarios de modèle de présentation : il existe un petit nombre de présentations connues qui servent de pools de diapositives qui peuvent être insérées par le complément.</span><span class="sxs-lookup"><span data-stu-id="b92e3-107">The slide insertion APIs are primarily used in presentation template scenarios: There are a small number of known presentations which serve as pools of slides that can be inserted by the add-in.</span></span> <span data-ttu-id="b92e3-108">Dans ce cas, vous ou le client devez créer et gérer une source de données qui met en corrélation le critère de sélection (comme les titres ou les images des diapositives) et les ID de diapositive.</span><span class="sxs-lookup"><span data-stu-id="b92e3-108">In such a scenario, either you or the customer must create and maintain a data source that correlates the selection criterion (such as slide titles or images) with slide IDs.</span></span> <span data-ttu-id="b92e3-109">Les API peuvent également être utilisées dans les scénarios dans lesquels l’utilisateur peut insérer des diapositives à partir de n’importe quelle présentation arbitraire, mais dans ce scénario, l’utilisateur est limité à insérer *toutes* les diapositives de la présentation source.</span><span class="sxs-lookup"><span data-stu-id="b92e3-109">The APIs can also be used in scenarios where the user can insert slides from any arbitrary presentation, but in that scenario the user is effectively limited to inserting *all* the slides from the source presentation.</span></span> <span data-ttu-id="b92e3-110">Pour plus d’informations à ce sujet, voir [sélection des diapositives à insérer](#selecting-which-slides-to-insert) .</span><span class="sxs-lookup"><span data-stu-id="b92e3-110">See [Selecting which slides to insert](#selecting-which-slides-to-insert) for more information about this.</span></span>

<span data-ttu-id="b92e3-111">Il existe deux étapes pour insérer des diapositives d’une présentation dans une autre.</span><span class="sxs-lookup"><span data-stu-id="b92e3-111">There are two steps to inserting slides from one presentation into another.</span></span>

1. <span data-ttu-id="b92e3-112">Convertissez le fichier de présentation source (. pptx) en une chaîne au format Base64.</span><span class="sxs-lookup"><span data-stu-id="b92e3-112">Convert the source presentation file (.pptx) into a base64-formatted string.</span></span>
1. <span data-ttu-id="b92e3-113">Utilisez la `insertSlidesFromBase64` méthode pour insérer une ou plusieurs diapositives à partir du fichier Base64 dans la présentation active.</span><span class="sxs-lookup"><span data-stu-id="b92e3-113">Use the `insertSlidesFromBase64` method to insert one or more slides from the base64 file into the current presentation.</span></span>

## <a name="convert-the-source-presentation-to-base64"></a><span data-ttu-id="b92e3-114">Convertir la présentation source en base64</span><span class="sxs-lookup"><span data-stu-id="b92e3-114">Convert the source presentation to base64</span></span>

<span data-ttu-id="b92e3-115">Il existe plusieurs façons de convertir un fichier en base64.</span><span class="sxs-lookup"><span data-stu-id="b92e3-115">There are many ways to convert a file to base64.</span></span> <span data-ttu-id="b92e3-116">Le langage de programmation et la bibliothèque que vous utilisez et s’il faut effectuer une conversion côté serveur de votre complément ou côté client est déterminé par votre scénario.</span><span class="sxs-lookup"><span data-stu-id="b92e3-116">Which programming language and library you use, and whether to convert on the server-side of your add-in or the client-side is determined by your scenario.</span></span> <span data-ttu-id="b92e3-117">En règle générale, vous effectuerez la conversion en JavaScript côté client à l’aide d’un objet [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) .</span><span class="sxs-lookup"><span data-stu-id="b92e3-117">Most commonly, you'll do the conversion in JavaScript on the client-side by using a [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) object.</span></span> <span data-ttu-id="b92e3-118">L’exemple suivant illustre cette pratique.</span><span class="sxs-lookup"><span data-stu-id="b92e3-118">The following example shows this practice.</span></span>

1. <span data-ttu-id="b92e3-119">Commencez par obtenir une référence au fichier PowerPoint source.</span><span class="sxs-lookup"><span data-stu-id="b92e3-119">Begin by getting a reference to the source PowerPoint file.</span></span> <span data-ttu-id="b92e3-120">Dans cet exemple, nous allons utiliser un `<input>` contrôle de type `file` pour inviter l’utilisateur à choisir un fichier.</span><span class="sxs-lookup"><span data-stu-id="b92e3-120">In this example, we will use an `<input>` control of type `file` to prompt the user to choose a file.</span></span> <span data-ttu-id="b92e3-121">Ajoutez le balisage suivant à la page de complément.</span><span class="sxs-lookup"><span data-stu-id="b92e3-121">Add the following markup to the add-in page.</span></span>

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    <span data-ttu-id="b92e3-122">Ce balisage ajoute l’interface utilisateur dans la capture d’écran suivante à la page :</span><span class="sxs-lookup"><span data-stu-id="b92e3-122">This markup adds the UI in the following screenshot to the page:</span></span>

    ![Capture d’écran illustrant un contrôle d’entrée de type de fichier HTML précédé d’une phrase pédagogique en lisant « sélectionnez une présentation PowerPoint à partir de laquelle insérer des diapositives ».](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > <span data-ttu-id="b92e3-125">Il existe de nombreuses autres façons d’obtenir un fichier PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="b92e3-125">There are many other ways to get a PowerPoint file.</span></span> <span data-ttu-id="b92e3-126">Par exemple, si le fichier est stocké sur OneDrive ou SharePoint, vous pouvez utiliser Microsoft Graph pour le télécharger.</span><span class="sxs-lookup"><span data-stu-id="b92e3-126">For example, if the file is stored on OneDrive or SharePoint, you can use Microsoft Graph to download it.</span></span> <span data-ttu-id="b92e3-127">Pour plus d’informations, consultez la rubrique [utilisation de fichiers dans Microsoft Graph](/graph/api/resources/onedrive) et [accès à des fichiers avec Microsoft Graph](/learn/modules/msgraph-access-file-data/).</span><span class="sxs-lookup"><span data-stu-id="b92e3-127">For more information, see [Working with files in Microsoft Graph](/graph/api/resources/onedrive) and [Access Files with Microsoft Graph](/learn/modules/msgraph-access-file-data/).</span></span>

2. <span data-ttu-id="b92e3-128">Ajoutez le code suivant au JavaScript du complément pour assigner une fonction à l’événement du contrôle d’entrée `change` .</span><span class="sxs-lookup"><span data-stu-id="b92e3-128">Add the following code to the add-in's JavaScript to assign a function to the input control's `change` event.</span></span> <span data-ttu-id="b92e3-129">(Vous créez la `storeFileAsBase64` fonction à l’étape suivante.)</span><span class="sxs-lookup"><span data-stu-id="b92e3-129">(You create the `storeFileAsBase64` function in the next step.)</span></span>

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. <span data-ttu-id="b92e3-130">Ajoutez le code suivant.</span><span class="sxs-lookup"><span data-stu-id="b92e3-130">Add the following code.</span></span> <span data-ttu-id="b92e3-131">Notez ce qui suit à propos de ce code :</span><span class="sxs-lookup"><span data-stu-id="b92e3-131">Note the following about this code,:</span></span>

    - <span data-ttu-id="b92e3-132">La `reader.readAsDataURL` méthode convertit le fichier en base64 et le stocke dans la `reader.result` propriété.</span><span class="sxs-lookup"><span data-stu-id="b92e3-132">The `reader.readAsDataURL` method converts the file to base64 and stores it in the `reader.result` property.</span></span> <span data-ttu-id="b92e3-133">Une fois la méthode terminée, le gestionnaire d’événements est déclenché `onload` .</span><span class="sxs-lookup"><span data-stu-id="b92e3-133">When the method completes, it triggers the `onload` event handler.</span></span>
    - <span data-ttu-id="b92e3-134">Le `onload` Gestionnaire d’événements supprime les métadonnées du fichier encodé et stocke la chaîne encodée dans une variable globale.</span><span class="sxs-lookup"><span data-stu-id="b92e3-134">The `onload` event handler trims metadata off of the encoded file and stores the encoded string in a global variable.</span></span>
    - <span data-ttu-id="b92e3-135">La chaîne codée en base64 est stockée globalement, car elle sera lue par une autre fonction que vous créez dans une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="b92e3-135">The base64-encoded string is stored globally because it will be read by another function that you create in a later step.</span></span>

    ```javascript
    let chosenFileBase64;

    async function storeFileAsBase64() {
        const reader = new FileReader();

        reader.onload = async (event) => {
            const startIndex = reader.result.toString().indexOf("base64,");
            const copyBase64 = reader.result.toString().substr(startIndex + 7);

            chosenFileBase64 = copyBase64;
        };

        const myFile = document.getElementById("file") as HTMLInputElement;
        reader.readAsDataURL(myFile.files[0]);
    }
    ```

## <a name="insert-slides-with-insertslidesfrombase64"></a><span data-ttu-id="b92e3-136">Insérer des diapositives avec insertSlidesFromBase64</span><span class="sxs-lookup"><span data-stu-id="b92e3-136">Insert slides with insertSlidesFromBase64</span></span>

<span data-ttu-id="b92e3-137">Votre complément insère des diapositives d’une autre présentation PowerPoint dans la présentation actuelle à l’aide de la méthode [Presentation. insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) .</span><span class="sxs-lookup"><span data-stu-id="b92e3-137">Your add-in inserts slides from another PowerPoint presentation into the current presentation with the [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) method.</span></span> <span data-ttu-id="b92e3-138">Voici un exemple simple dans lequel toutes les diapositives de la présentation source sont insérées au début de la présentation en cours et les diapositives insérées conservent la mise en forme du fichier source.</span><span class="sxs-lookup"><span data-stu-id="b92e3-138">The following is a simple example in which all of the slides from the source presentation are inserted at the beginning of the current presentation and the inserted slides keep the formatting of the source file.</span></span> <span data-ttu-id="b92e3-139">Notez qu' `chosenFileBase64` il s’agit d’une variable globale qui contient une version codée en base64 d’un fichier de présentation PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="b92e3-139">Note that `chosenFileBase64` is a global variable that holds a base64-encoded version of a PowerPoint presentation file.</span></span>

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

<span data-ttu-id="b92e3-140">Vous pouvez contrôler certains aspects du résultat de l’insertion, y compris où les diapositives sont insérées et déterminer si elles obtiennent la mise en forme source ou cible, en transmettant un objet [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) comme deuxième paramètre à `insertSlidesFromBase64` .</span><span class="sxs-lookup"><span data-stu-id="b92e3-140">You can control some aspects of the insertion result, including where the slides are inserted and whether they get the source or target formatting , by passing an [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) object as a second parameter to `insertSlidesFromBase64`.</span></span> <span data-ttu-id="b92e3-141">Voici un exemple.</span><span class="sxs-lookup"><span data-stu-id="b92e3-141">The following is an example.</span></span> <span data-ttu-id="b92e3-142">Tenez compte du code suivant :</span><span class="sxs-lookup"><span data-stu-id="b92e3-142">About this code, note:</span></span>

- <span data-ttu-id="b92e3-143">Il existe deux valeurs possibles pour la `formatting` propriété : « UseDestinationTheme » et « KeepSourceFormatting ».</span><span class="sxs-lookup"><span data-stu-id="b92e3-143">There are two possible values for the `formatting` property: "UseDestinationTheme" and "KeepSourceFormatting".</span></span> <span data-ttu-id="b92e3-144">Vous pouvez également utiliser l' `InsertSlideFormatting` énumération (par exemple, `PowerPoint.InsertSlideFormatting.useDestinationTheme` ).</span><span class="sxs-lookup"><span data-stu-id="b92e3-144">Optionally, you can use the `InsertSlideFormatting` enum, (e.g., `PowerPoint.InsertSlideFormatting.useDestinationTheme`).</span></span>
- <span data-ttu-id="b92e3-145">La fonction insère les diapositives de la présentation source immédiatement après la diapositive spécifiée par la `targetSlideId` propriété.</span><span class="sxs-lookup"><span data-stu-id="b92e3-145">The function will insert the slides from the source presentation immediately after the slide specified by the `targetSlideId` property.</span></span> <span data-ttu-id="b92e3-146">La valeur de cette propriété est une chaîne d’une des trois formes possibles : \***nnn \* #**, \* *#* mmmmmmmmm \* \* \* ou \**_nnn_ #* mmmmmmmmm \* \* \*, où *nnn* est l’ID de la diapositive (généralement 3 chiffres) et *mmmmmmmmm* est l’ID de création de la diapositive (généralement 9 chiffres).</span><span class="sxs-lookup"><span data-stu-id="b92e3-146">The value of this property is a string of one of three possible forms: \***nnn\*#**, \**#* mmmmmmmmm\*\*\*, or \**_nnn_#* mmmmmmmmm\*\*\*, where *nnn* is the slide's ID (typically 3 digits) and *mmmmmmmmm* is the slide's creation ID (typically 9 digits).</span></span> <span data-ttu-id="b92e3-147">Voici quelques exemples :, `267#763315295` `267#` et `#763315295` .</span><span class="sxs-lookup"><span data-stu-id="b92e3-147">Some examples are `267#763315295`, `267#`, and `#763315295`.</span></span>

```javascript
async function insertSlidesDestinationFormatting() {
  await PowerPoint.run(async function(context) {
    context.presentation
    .insertSlidesFromBase64(chosenFileBase64,
                            {
                                formatting: "UseDestinationTheme",
                                targetSlideId: "267#"
                            }
                          );
    await context.sync();
  });
}
```

<span data-ttu-id="b92e3-148">Bien entendu, vous ne saurez généralement pas au moment du code l’ID ou l’ID de création de la diapositive cible.</span><span class="sxs-lookup"><span data-stu-id="b92e3-148">Of course, you typically won't know at coding time the ID or creation ID of the target slide.</span></span> <span data-ttu-id="b92e3-149">Plus communément, un complément demande aux utilisateurs de sélectionner la diapositive cible.</span><span class="sxs-lookup"><span data-stu-id="b92e3-149">More commonly, an add-in will ask users to select the target slide.</span></span> <span data-ttu-id="b92e3-150">Les étapes suivantes montrent comment obtenir l’ID \***nnn \* #** de la diapositive actuellement sélectionnée et l’utiliser comme diapositive cible.</span><span class="sxs-lookup"><span data-stu-id="b92e3-150">The following steps show how to get the \***nnn\*#** ID of the currently selected slide and use it as the target slide.</span></span>

1. <span data-ttu-id="b92e3-151">Créez une fonction qui obtient l’ID de la diapositive actuellement sélectionnée à l’aide de la [Office.context.docméthode ument. getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) des API JavaScript communes.</span><span class="sxs-lookup"><span data-stu-id="b92e3-151">Create a function that gets the ID of the currently selected slide by using the [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) method of the Common JavaScript APIs.</span></span> <span data-ttu-id="b92e3-152">Voici un exemple.</span><span class="sxs-lookup"><span data-stu-id="b92e3-152">The following is an example.</span></span> <span data-ttu-id="b92e3-153">Notez que l’appel à `getSelectedDataAsync` est incorporé dans une fonction de retour à la vente.</span><span class="sxs-lookup"><span data-stu-id="b92e3-153">Note that the call to `getSelectedDataAsync` is embedded in a Promise-returning function.</span></span> <span data-ttu-id="b92e3-154">Pour plus d’informations sur les raisons et la procédure à suivre, consultez [la rubrique Wrap Common-APIs dans les fonctions de retour à la vente](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span><span class="sxs-lookup"><span data-stu-id="b92e3-154">For more information about why and how to do this, see [Wrap Common-APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span></span>

 
    ```javascript
    function getSelectedSlideID() {
      return new OfficeExtension.Promise<string>(function (resolve, reject) {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
          try {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              reject(console.error(asyncResult.error.message));
            } else {
              resolve(asyncResult.value.slides[0].id);
            }
          }
          catch (error) {
            reject(console.log(error));
          }
        });
      })
    }
    ```

1. <span data-ttu-id="b92e3-155">Appelez votre nouvelle fonction à l’intérieur de [PowerPoint. Run ()](/javascript/api/powerpoint#PowerPoint_run_batch_) de la fonction main et transmettez l’ID qu’elle renvoie (concaténé avec le symbole « # ») en tant que valeur de la `targetSlideId` propriété du `InsertSlideOptions` paramètre.</span><span class="sxs-lookup"><span data-stu-id="b92e3-155">Call your new function inside the [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) of the main function and pass the ID that it returns (concatenated with the "#" symbol) as the value of the `targetSlideId` property of the `InsertSlideOptions` parameter.</span></span> <span data-ttu-id="b92e3-156">Voici un exemple.</span><span class="sxs-lookup"><span data-stu-id="b92e3-156">The following is an example.</span></span>

    ```javascript
    async function insertAfterSelectedSlide() {
        await PowerPoint.run(async function(context) {

            const selectedSlideID = await getSelectedSlideID();

            context.presentation.insertSlidesFromBase64(chosenFileBase64, {
                formatting: "UseDestinationTheme",
                targetSlideId: selectedSlideID + "#"
            });

            await context.sync();
        });
    }
    ```

### <a name="selecting-which-slides-to-insert"></a><span data-ttu-id="b92e3-157">Sélection des diapositives à insérer</span><span class="sxs-lookup"><span data-stu-id="b92e3-157">Selecting which slides to insert</span></span>

<span data-ttu-id="b92e3-158">Vous pouvez également utiliser le paramètre [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) pour contrôler les diapositives de la présentation source qui doivent être insérées.</span><span class="sxs-lookup"><span data-stu-id="b92e3-158">You can also use the [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) parameter to control which slides from the source presentation are inserted.</span></span> <span data-ttu-id="b92e3-159">Pour ce faire, affectez un tableau des ID de diapositives de la présentation source à la `sourceSlideIds` propriété.</span><span class="sxs-lookup"><span data-stu-id="b92e3-159">You do this by assigning an array of the source presentation's slide IDs to the `sourceSlideIds` property.</span></span> <span data-ttu-id="b92e3-160">Voici un exemple qui insère quatre diapositives.</span><span class="sxs-lookup"><span data-stu-id="b92e3-160">The following is an example that inserts four slides.</span></span> <span data-ttu-id="b92e3-161">Notez que chaque chaîne dans le tableau doit respecter un ou l’autre des modèles utilisés pour la `targetSlideId` propriété.</span><span class="sxs-lookup"><span data-stu-id="b92e3-161">Note that each string in the array must follow one or another of the patterns used for the `targetSlideId` property.</span></span>

```javascript
async function insertAfterSelectedSlide() {
    await PowerPoint.run(async function(context) {
        const selectedSlideID = await getSelectedSlideID();
        context.presentation.insertSlidesFromBase64(chosenFileBase64, {
            formatting: "UseDestinationTheme",
            targetSlideId: selectedSlideID + "#",
            sourceSlideIds: ["267#763315295", "256#", "#926310875", "1270#"]
        });

        await context.sync();
    });
}
```

> [!NOTE]
> <span data-ttu-id="b92e3-162">Les diapositives sont insérées dans le même ordre relatif que celui dans lequel elles apparaissent dans la présentation source, quel que soit l’ordre dans lequel elles apparaissent dans le tableau.</span><span class="sxs-lookup"><span data-stu-id="b92e3-162">The slides will be inserted in the same relative order in which they appear in the source presentation, regardless of the order in which they appear in the array.</span></span>

<span data-ttu-id="b92e3-163">Il n’existe aucun moyen pratique pour les utilisateurs de découvrir l’ID ou l’ID de création d’une diapositive dans la présentation source.</span><span class="sxs-lookup"><span data-stu-id="b92e3-163">There is no practical way that users can discover the ID or creation ID of a slide in the source presentation.</span></span> <span data-ttu-id="b92e3-164">Pour cette raison, vous pouvez uniquement utiliser la `sourceSlideIds` propriété lorsque vous avez identifié les ID de source au moment du codage ou que votre complément peut les récupérer lors de l’exécution à partir d’une source de données.</span><span class="sxs-lookup"><span data-stu-id="b92e3-164">For this reason, you can really only use the `sourceSlideIds` property when either you know the source IDs at coding time or your add-in can retrieve them at runtime from some data source.</span></span> <span data-ttu-id="b92e3-165">Étant donné que les utilisateurs ne peuvent pas mémoriser les ID de diapositive, vous avez également besoin d’un moyen pour permettre à l’utilisateur de sélectionner des diapositives, par exemple par titre ou par image, puis de corréler chaque titre ou image avec l’ID de la diapositive.</span><span class="sxs-lookup"><span data-stu-id="b92e3-165">Because users cannot be expected to memorize slide IDs, you also need a way to enable the user to select slides, perhaps by title or by an image, and then correlate each title or image with the slide's ID.</span></span>

<span data-ttu-id="b92e3-166">En conséquence, la `sourceSlideIds` propriété est principalement utilisée dans les scénarios de modèle de présentation : le complément est conçu pour fonctionner avec un ensemble spécifique de présentations qui servent de pools de diapositives qui peuvent être insérées.</span><span class="sxs-lookup"><span data-stu-id="b92e3-166">Accordingly, the `sourceSlideIds` property is primarily used in presentation template scenarios: The add-in is designed to work with a specific set of presentations that serve as pools of slides that can be inserted.</span></span> <span data-ttu-id="b92e3-167">Dans ce cas, vous ou le client devez créer et gérer une source de données qui met en corrélation un critère de sélection (tel que des titres ou des images) avec des ID de diapositive ou des ID de création de diapositives qui ont été créés à partir de l’ensemble de présentations source possibles.</span><span class="sxs-lookup"><span data-stu-id="b92e3-167">In such a scenario, either you or the customer must create and maintain a data source that correlates a selection criterion (such as titles or images) with slide IDs or slide creation IDs that has been constructed from the set of possible source presentations.</span></span>

## <a name="delete-slides"></a><span data-ttu-id="b92e3-168">Supprimer des diapositives</span><span class="sxs-lookup"><span data-stu-id="b92e3-168">Delete slides</span></span>

<span data-ttu-id="b92e3-169">Vous pouvez supprimer une diapositive en obtenant une référence à l’objet [Slide](/javascript/api/powerpoint/powerpoint.slide) qui représente la diapositive et appeler la `Slide.delete` méthode.</span><span class="sxs-lookup"><span data-stu-id="b92e3-169">You can delete a slide by getting a reference to the [Slide](/javascript/api/powerpoint/powerpoint.slide) object that represents the slide and call the `Slide.delete` method.</span></span> <span data-ttu-id="b92e3-170">Voici un exemple dans lequel la quatrième diapositive est supprimée.</span><span class="sxs-lookup"><span data-stu-id="b92e3-170">The following is an example in which the 4th slide is deleted.</span></span>

```javascript
async function deleteSlide() {
  await PowerPoint.run(async function(context) {

    // The slide index is zero-based. 
    const slide = context.presentation.slides.getItemAt(3);
    slide.delete();
    await context.sync();
  });
}
```
