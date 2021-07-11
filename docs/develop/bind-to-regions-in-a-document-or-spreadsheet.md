---
title: Lier des régions dans un document ou une feuille de calcul
description: Découvrez comment utiliser la liaison pour garantir un accès cohérent à une région ou à un élément spécifique d’un document ou d’une feuille de calcul via un identificateur.
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: c9a658653c562de446f3b8e5f1ea192ddfcf3b21
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350000"
---
# <a name="bind-to-regions-in-a-document-or-spreadsheet"></a><span data-ttu-id="da664-103">Lier des régions dans un document ou une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="da664-103">Bind to regions in a document or spreadsheet</span></span>

<span data-ttu-id="da664-104">L’accès aux données basées sur une liaison permet aux compléments de contenu et du volet Office d’accéder de façon cohérente à une zone particulière d’un document ou d’une feuille de calcul au moyen d’un identificateur.</span><span class="sxs-lookup"><span data-stu-id="da664-104">Binding-based data access enables content and task pane add-ins to consistently access a particular region of a document or spreadsheet through an identifier.</span></span> <span data-ttu-id="da664-105">Le complément doit d’abord établir la liaison en appelant l’une des méthodes qui associent une partie du document à un identificateur unique : [addFromPromptAsync], [addFromSelectionAsync] ou [addFromNamedItemAsync].</span><span class="sxs-lookup"><span data-stu-id="da664-105">The add-in first needs to establish the binding by calling one of the methods that associates a portion of the document with a unique identifier: [addFromPromptAsync], [addFromSelectionAsync], or [addFromNamedItemAsync].</span></span> <span data-ttu-id="da664-106">Une fois la liaison établie, le complément peut utiliser l’identificateur fourni pour accéder aux données contenues dans la région associée du document ou de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="da664-106">After the binding is established, the add-in can use the provided identifier to access the data contained in the associated region of the document or spreadsheet.</span></span> <span data-ttu-id="da664-107">La création de liaisons fournit la valeur suivante à votre add-in.</span><span class="sxs-lookup"><span data-stu-id="da664-107">Creating bindings provides the following value to your add-in.</span></span>

- <span data-ttu-id="da664-108">Elle permet l’accès aux structures de données communes sur les applications Office prises en charge, telles que : tableaux, plages ou texte (série contiguë de caractères).</span><span class="sxs-lookup"><span data-stu-id="da664-108">Permits access to common data structures across supported Office applications, such as: tables, ranges, or text (a contiguous run of characters).</span></span>
- <span data-ttu-id="da664-109">Elle permet les opérations de lecture/écriture sans exiger que l’utilisateur effectue une sélection.</span><span class="sxs-lookup"><span data-stu-id="da664-109">Enables read/write operations without requiring the user to make a selection.</span></span>
- <span data-ttu-id="da664-p102">Elle établit une relation entre le complément et les données du document. Les liaisons persistent dans le document et sont accessibles par la suite.</span><span class="sxs-lookup"><span data-stu-id="da664-p102">Establishes a relationship between the add-in and the data in the document. Bindings are persisted in the document, and can be accessed at a later time.</span></span>

<span data-ttu-id="da664-p103">L’établissement d’une liaison vous permet également de vous abonner aux données et aux événements de changement de sélection qui sont concernés par cette région particulière du document ou de la feuille de calcul. Cela signifie que le complément est seulement notifié des changements qui surviennent dans la région délimitée, par opposition aux changements généraux affectant l’ensemble du document ou de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="da664-p103">Establishing a binding also allows you to subscribe to data and selection change events that are scoped to that particular region of the document or spreadsheet. This means that the add-in is only notified of changes that happen within the bound region as opposed to general changes across the whole document or spreadsheet.</span></span>

<span data-ttu-id="da664-p104">L’objet [Bindings] expose une méthode [getAllAsync] qui donne accès à toutes les liaisons établies dans le document ou la feuille de calcul. Une liaison individuelle est accessible par son ID à l’aide de la méthode [Bindings.getBindingByIdAsync] ou [Office.select]. Vous pouvez établir de nouvelles liaisons et supprimer des liaisons existantes en utilisant l’une des méthodes suivantes de l’objet [Bindings] : [addFromSelectionAsync], [addFromPromptAsync], [addFromNamedItemAsync] ou [releaseByIdAsync].</span><span class="sxs-lookup"><span data-stu-id="da664-p104">The [Bindings] object exposes a [getAllAsync] method that gives access to the set of all bindings established on the document or spreadsheet. An individual binding can be accessed by its ID using either the Bindings.[getByIdAsync] or [Office.select] methods. You can establish new bindings as well as remove existing ones by using one of the following methods of the [Bindings] object: [addFromSelectionAsync], [addFromPromptAsync], [addFromNamedItemAsync], or [releaseByIdAsync].</span></span>

## <a name="binding-types"></a><span data-ttu-id="da664-117">Types de liaison</span><span class="sxs-lookup"><span data-stu-id="da664-117">Binding types</span></span>

<span data-ttu-id="da664-118">Il existe [trois types de liaisons][Office. BindingType] que vous spécifiez avec le paramètre _bindingType_ lorsque vous créez une liaison avec les méthodes [addFromSelectionAsync], [addFromPromptAsync] ou [addFromNamedItemAsync] :</span><span class="sxs-lookup"><span data-stu-id="da664-118">There are [three different types of bindings][Office.BindingType] that you specify with the _bindingType_ parameter when you create a binding with the [addFromSelectionAsync], [addFromPromptAsync] or [addFromNamedItemAsync] methods:</span></span>

1. <span data-ttu-id="da664-119">**[Liaison de texte][TextBinding]** - Établit une liaison à une zone du document qui est représentée en tant que texte.</span><span class="sxs-lookup"><span data-stu-id="da664-119">**[Text Binding][TextBinding]** - Binds to a region of the document that can be represented as text.</span></span>

    <span data-ttu-id="da664-p105">Dans Word, la plupart des sélections contiguës sont valides, tandis que dans Excel, seules les sélections de cellules uniques peuvent être la cible d’une liaison de texte. Dans Excel, seul le texte brut est pris en charge. Dans Word, trois formats sont pris en charge : texte brut, HTML et Open XML pour Office.</span><span class="sxs-lookup"><span data-stu-id="da664-p105">In Word, most contiguous selections are valid, while in Excel only single cell selections can be the target of a text binding. In Excel, only plain text is supported. In Word, three formats are supported: plain text, HTML, and Open XML for Office.</span></span>

2. <span data-ttu-id="da664-123">**[Matrix Binding][MatrixBinding]** : se lie à une région fixe d’un document qui contient des données tabulaires sans en-têtes. Les données d’une liaison de matrice sont écrites ou lues sous la forme d’un tableau à deux **dimensions,** qui en JavaScript est implémenté sous la forme d’un tableau de tableaux.</span><span class="sxs-lookup"><span data-stu-id="da664-123">**[Matrix Binding][MatrixBinding]** - Binds to a fixed region of a document that contains tabular data without headers.Data in a matrix binding is written or read as a two dimensional **Array**, which in JavaScript is implemented as an array of arrays.</span></span> <span data-ttu-id="da664-124">Par exemple, deux lignes de valeurs **string** dans deux colonnes peuvent être écrites ou lues comme `[['a', 'b'], ['c', 'd']]`, et une colonne unique de trois lignes peut être écrite ou lue comme `[['a'], ['b'], ['c']]`.</span><span class="sxs-lookup"><span data-stu-id="da664-124">For example, two rows of **string** values in two columns can be written or read as `[['a', 'b'], ['c', 'd']]`, and a single column of three rows can be written or read as `[['a'], ['b'], ['c']]`.</span></span>

    <span data-ttu-id="da664-p107">Dans Excel, toute sélection contiguë de cellules peut être utilisée pour établir une liaison de matrice. Dans Word, seuls les tableaux prennent en charge la liaison de matrice.</span><span class="sxs-lookup"><span data-stu-id="da664-p107">In Excel, any contiguous selection of cells can be used to establish a matrix binding. In Word, only tables support matrix binding.</span></span>

3. <span data-ttu-id="da664-p108">**[Liaison de tableau][TableBinding]** - Établit une liaison à une zone d’un document qui contient un tableau avec des en-têtes. Les données dans une liaison de tableau sont écrites ou lues comme un objet [TableData](/javascript/api/office/office.tabledata). L’objet `TableData` expose les données via les propriétés `headers` et `rows`.</span><span class="sxs-lookup"><span data-stu-id="da664-p108">**[Table Binding][TableBinding]** - Binds to a region of a document that contains a table with headers.Data in a table binding is written or read as a [TableData](/javascript/api/office/office.tabledata) object. The `TableData` object exposes the data through the `headers` and `rows` properties.</span></span>

    <span data-ttu-id="da664-p109">Tout tableau Excel ou Word peut être la base d’une liaison de tableau. Une fois que vous établissez une liaison de tableau, chaque nouvelle ligne ou colonne qu’un utilisateur ajoute au tableau est automatiquement incluse dans la liaison.</span><span class="sxs-lookup"><span data-stu-id="da664-p109">Any Excel or Word table can be the basis for a table binding. After you establish a table binding, each new row or column a user adds to the table is automatically included in the binding.</span></span>

<span data-ttu-id="da664-131">Après la création d’une liaison à l’aide de l’une des trois méthodes « addFrom » de l’objet, vous pouvez utiliser les données et propriétés de la liaison à l’aide des méthodes de `Bindings` l’objet correspondant : [MatrixBinding,] [TableBinding]ou [TextBinding].</span><span class="sxs-lookup"><span data-stu-id="da664-131">After a binding is created by using one of the three "addFrom" methods of the `Bindings` object, you can work with the binding's data and properties by using the methods of the corresponding object: [MatrixBinding], [TableBinding], or [TextBinding].</span></span> <span data-ttu-id="da664-132">Ces trois objets héritent des méthodes [getDataAsync] et [setDataAsync] de l’objet `Binding` qui vous permettent d’interagir avec les données liées.</span><span class="sxs-lookup"><span data-stu-id="da664-132">All three of these objects inherit the [getDataAsync] and [setDataAsync] methods of the `Binding` object that enable you to interact with the bound data.</span></span>

> [!NOTE]
> <span data-ttu-id="da664-p111">**Quand devez-vous utiliser une liaison de matrice ou une liaison de tableau ?** Lorsque les données tabulaires avec lesquelles vous travaillez contiennent une ligne de total, vous devez utiliser une liaison de matrice si le script de votre complément doit accéder aux valeurs figurant dans la ligne de total ou détecter que la sélection de l’utilisateur figure dans la ligne de total. Si vous établissez une liaison de tableau pour des données tabulaires qui contiennent une ligne de total, la propriété [TableBinding.rowCount] et les propriétés `rowCount` et `startRow` de l’objet [BindingSelectionChangedEventArgs] dans les gestionnaires d’événements ne reflèteront pas la ligne de total dans leurs valeurs. Pour contourner cette limitation, vous devez établir une liaison de matrice pour travailler avec la ligne de total.</span><span class="sxs-lookup"><span data-stu-id="da664-p111">**When should you use matrix versus table bindings?** When the tabular data you are working with contains a total row, you must use a matrix binding if your add-in's script needs to access values in the total row or detect that the user's selection is in the total row. If you establish a table binding for tabular data that contains a total row, the [TableBinding.rowCount] property and the `rowCount` and `startRow` properties of the [BindingSelectionChangedEventArgs] object in event handlers won't reflect the total row in their values. To work around this limitation, you must use establish a matrix binding to work with the total row.</span></span>

## <a name="add-a-binding-to-the-users-current-selection"></a><span data-ttu-id="da664-137">Ajout d’une liaison à la sélection actuelle de l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="da664-137">Add a binding to the user's current selection</span></span>

<span data-ttu-id="da664-138">L’exemple suivant montre comment ajouter une liaison de texte nommée `myBinding` à la sélection actuelle dans un document à l’aide de la méthode [addFromSelectionAsync].</span><span class="sxs-lookup"><span data-stu-id="da664-138">The following example shows how to add a text binding called  `myBinding` to the current selection in a document by using the [addFromSelectionAsync] method.</span></span>

```js
Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="da664-p112">Dans cet exemple, le type de liaison spécifié est « Text ». Cela signifie qu’un objet [TextBinding] sera créé pour la sélection. Différents types de liaison exposent différentes données et opérations. [Office.BindingType] est une énumération des valeurs de types de liaison disponibles.</span><span class="sxs-lookup"><span data-stu-id="da664-p112">In this example, the specified binding type is text. This means that a [TextBinding] will be created for the selection. Different binding types expose different data and operations. [Office.BindingType] is an enumeration of available binding type values.</span></span>

<span data-ttu-id="da664-p113">Le deuxième paramètre facultatif est un objet qui spécifie l’ID de la nouvelle liaison créée. Si un ID n’est pas spécifié, un ID est généré automatiquement.</span><span class="sxs-lookup"><span data-stu-id="da664-p113">The second optional parameter is an object that specifies the ID of the new binding being created. If an ID is not specified, one is generated automatically.</span></span>

<span data-ttu-id="da664-p114">La fonction anonyme qui est passée dans la fonction comme paramètre final _callback_ est exécutée lorsque la création de la liaison est terminée. La fonction est appelée avec un seul paramètre, `asyncResult`, ce qui donne accès à un objet [AsyncResult] qui fournit l’état de l’appel. La propriété `AsyncResult.value` contient une référence à un objet [Binding] du type spécifié pour la liaison créée récemment. Vous pouvez utiliser cet objet [Binding] pour obtenir et définir les données.</span><span class="sxs-lookup"><span data-stu-id="da664-p114">The anonymous function that is passed into the function as the final  _callback_ parameter is executed when the creation of the binding is complete. The function is called with a single parameter, `asyncResult`, which provides access to an [AsyncResult] object that provides the status of the call. The `AsyncResult.value` property contains a reference to a [Binding] object of the type that is specified for the newly created binding. You can use this [Binding] object to get and set data.</span></span>

## <a name="add-a-binding-from-a-prompt"></a><span data-ttu-id="da664-149">Ajout d’une liaison à partir d’une invite</span><span class="sxs-lookup"><span data-stu-id="da664-149">Add a binding from a prompt</span></span>

<span data-ttu-id="da664-p115">L’exemple suivant indique comment ajouter une liaison de texte appelée `myBinding` à l’aide de la méthode [addFromPromptAsync]. Cette méthode permet à l’utilisateur de spécifier la plage pour la liaison à l’aide de l’invite de sélection de plage intégrée.</span><span class="sxs-lookup"><span data-stu-id="da664-p115">The following example shows how to add a text binding called  `myBinding` by using the [addFromPromptAsync] method. This method lets the user specify the range for the binding by using the application's built-in range selection prompt.</span></span>

```js
function bindFromPrompt() {
    Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        } else {
            write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
        }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

<span data-ttu-id="da664-p116">Dans cet exemple, le type de liaison spécifié est « Text ». Cela signifie qu’un objet [TextBinding] sera créé pour la sélection que l’utilisateur spécifie dans l’invite.</span><span class="sxs-lookup"><span data-stu-id="da664-p116">In this example, the specified binding type is text. This means that a [TextBinding] will be created for the selection that the user specifies in the prompt.</span></span>

<span data-ttu-id="da664-p117">Le deuxième paramètre est un objet qui contient l’ID de la nouvelle liaison créée. Si un ID n’est pas spécifié, un ID est généré automatiquement.</span><span class="sxs-lookup"><span data-stu-id="da664-p117">The second parameter is an object that contains the ID of the new binding being created. If an ID is not specified, one is generated automatically.</span></span>

<span data-ttu-id="da664-156">La fonction anonyme passée dans la fonction en tant que _troisième_ paramètre de rappel est exécutée lorsque la création de la liaison est terminée.</span><span class="sxs-lookup"><span data-stu-id="da664-156">The anonymous function passed into the function as the third _callback_ parameter is executed when the creation of the binding is complete.</span></span> <span data-ttu-id="da664-157">Lorsque la fonction de rappel s’exécute, l’objet [AsyncResult] contient le statut de l’appel et la nouvelle liaison.</span><span class="sxs-lookup"><span data-stu-id="da664-157">When the callback function executes, the [AsyncResult] object contains the status of the call and the newly created binding.</span></span>

<span data-ttu-id="da664-158">La figure 1 montre l’invite de sélection de plage intégrée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="da664-158">Figure 1 shows the built-in range selection prompt in Excel.</span></span>

<span data-ttu-id="da664-159">*Figure 1. Interface utilisateur de sélection de données dans Excel*</span><span class="sxs-lookup"><span data-stu-id="da664-159">*Figure 1. Excel Select Data UI*</span></span>

![Capture d’écran montrant la boîte de dialogue Sélectionner des données.](../images/agave-api-overview-excel-selection-ui.png)

## <a name="add-a-binding-to-a-named-item"></a><span data-ttu-id="da664-161">Ajout d’une liaison à un élément nommé</span><span class="sxs-lookup"><span data-stu-id="da664-161">Add a binding to a named item</span></span>

<span data-ttu-id="da664-162">L’exemple suivant montre comment ajouter une liaison à l’élément nommé existant en tant que liaison « matrix » à l’aide de la méthode `myRange` [addFromNamedItemAsync,] et affecte la liaison en tant que `id` « myMatrix ».</span><span class="sxs-lookup"><span data-stu-id="da664-162">The following example shows how to add a binding to the existing `myRange` named item as a "matrix" binding by using the [addFromNamedItemAsync] method, and assigns the binding's `id` as "myMatrix".</span></span>

```js
function bindNamedItem() {
    Office.context.document.bindings.addFromNamedItemAsync("myRange", "matrix", {id:'myMatrix'}, function (result) {
        if (result.status == 'succeeded'){
            write('Added new binding with type: ' + result.value.type + ' and id: ' + result.value.id);
            }
        else
            write('Error: ' + result.error.message);
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```

<span data-ttu-id="da664-163">**Pour Excel**, le paramètre de la méthode `itemName` [addFromNamedItemAsync] peut faire référence à une plage nommée existante, à une plage spécifiée avec le style de référence ou à un `A1` `("A1:A3")` tableau.</span><span class="sxs-lookup"><span data-stu-id="da664-163">**For Excel**, the `itemName` parameter of the [addFromNamedItemAsync] method can refer to an existing named range, a range specified with the `A1` reference style `("A1:A3")`, or a table.</span></span> <span data-ttu-id="da664-164">Par défaut, l’ajout d’un tableau dans Excel entraîne l’affectation du nom « Tableau1 » pour le premier tableau que vous ajoutez, « Tableau2 » pour le deuxième tableau que vous ajoutez, et ainsi de suite.</span><span class="sxs-lookup"><span data-stu-id="da664-164">By default, adding a table in Excel assigns the name "Table1" for the first table you add, "Table2" for the second table you add, and so on.</span></span> <span data-ttu-id="da664-165">Pour attribuer un nom significatif à une table dans l’interface utilisateur Excel, utilisez la propriété dans la | `Table Name` **Onglet** Création du ruban.</span><span class="sxs-lookup"><span data-stu-id="da664-165">To assign a meaningful name for a table in the Excel UI, use the `Table Name` property on the **Table Tools | Design** tab of the ribbon.</span></span>

> [!NOTE]
> <span data-ttu-id="da664-166">Dans Excel, lorsque vous spécifiez un tableau en tant qu’élément nommé, vous devez qualifier complètement le nom pour inclure le nom de la feuille de calcul dans le nom du tableau au format suivant :`"Sheet1!Table1"`</span><span class="sxs-lookup"><span data-stu-id="da664-166">In Excel, when specifying a table as a named item, you must fully qualify the name to include the worksheet name in the name of the table in this format: `"Sheet1!Table1"`</span></span>

<span data-ttu-id="da664-167">L’exemple suivant crée une liaison dans Excel aux trois premières cellules de la colonne A ( ), affecte l’ID, puis écrit trois noms de ville dans `"A1:A3"` `"MyCities"` cette liaison.</span><span class="sxs-lookup"><span data-stu-id="da664-167">The following example creates a binding in Excel to the first three cells in column A ( `"A1:A3"`), assigns the id `"MyCities"`, and then writes three city names to that binding.</span></span>

```js
 function bindingFromA1Range() {
    Office.context.document.bindings.addFromNamedItemAsync("A1:A3", "matrix", {id: "MyCities" },
        function (asyncResult) {
            if (asyncResult.status == "failed") {
                write('Error: ' + asyncResult.error.message);
            }
            else {
                // Write data to the new binding.
                Office.select("bindings#MyCities").setDataAsync([['Berlin'], ['Munich'], ['Duisburg']], { coercionType: "matrix" },
                    function (asyncResult) {
                        if (asyncResult.status == "failed") {
                            write('Error: ' + asyncResult.error.message);
                        }
                    });
            }
        });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="da664-168">**Pour Word**, le paramètre de la méthode `itemName` [addFromNamedItemAsync] fait référence à la propriété `Title` d’un `Rich Text` contrôle de contenu.</span><span class="sxs-lookup"><span data-stu-id="da664-168">**For Word**, the `itemName` parameter of the [addFromNamedItemAsync] method refers to the `Title` property of a `Rich Text` content control.</span></span> <span data-ttu-id="da664-169">(Vous ne pouvez réaliser de liaison avec des contrôles de contenu différents du contrôle de contenu `Rich Text`.)</span><span class="sxs-lookup"><span data-stu-id="da664-169">(You can't bind to content controls other than the `Rich Text` content control.)</span></span>

<span data-ttu-id="da664-170">Par défaut, aucune valeur n’est affectée à un contrôle `Title*` de contenu.</span><span class="sxs-lookup"><span data-stu-id="da664-170">By default, a content control has no `Title*`value assigned.</span></span> <span data-ttu-id="da664-171">Pour affecter un nom significatif dans l’interface utilisateur de Word, après l’insertion d’un contrôle de contenu **Texte enrichi** à partir du groupe **Contrôles** sur l’onglet **Développeur** du ruban, utilisez la commande **Propriétés** du groupe **Contrôles** pour afficher la boîte de dialogue **Propriétés du contrôle de contenu**.</span><span class="sxs-lookup"><span data-stu-id="da664-171">To assign a meaningful name in the Word UI, after inserting a **Rich Text** content control from the **Controls** group on the **Developer** tab of the ribbon, use the **Properties** command in the **Controls** group to display the **Content Control Properties** dialog box.</span></span> <span data-ttu-id="da664-172">Définissez ensuite la propriété du contrôle de contenu sur le nom que `Title` vous souhaitez référencer à partir de votre code.</span><span class="sxs-lookup"><span data-stu-id="da664-172">Then set the `Title` property of the content control to the name you want to reference from your code.</span></span>

<span data-ttu-id="da664-173">L’exemple suivant crée une liaison de texte dans Word à un contrôle de contenu de texte enrichi nommé , affecte l’ID, puis affiche `"FirstName"`  `"firstName"` ces informations.</span><span class="sxs-lookup"><span data-stu-id="da664-173">The following example creates a text binding in Word to a rich text content control named `"FirstName"`, assigns the **id** `"firstName"`, and then displays that information.</span></span>

```js
function bindContentControl() {
    Office.context.document.bindings.addFromNamedItemAsync('FirstName', 
        Office.BindingType.Text, {id:'firstName'},
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                write('Control bound. Binding.id: '
                    + result.value.id + ' Binding.type: ' + result.value.type);
            } else {
                write('Error:', result.error.message);
            }
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

## <a name="get-all-bindings"></a><span data-ttu-id="da664-174">Obtention de toutes les liaisons</span><span class="sxs-lookup"><span data-stu-id="da664-174">Get all bindings</span></span>

<span data-ttu-id="da664-175">L’exemple suivant montre comment obtenir toutes les liaisons dans un document en utilisant la méthode Bindings.[getAllAsync].</span><span class="sxs-lookup"><span data-stu-id="da664-175">The following example shows how to get all bindings in a document by using the Bindings.[getAllAsync] method.</span></span>

```js
Office.context.document.bindings.getAllAsync(function (asyncResult) {
    var bindingString = '';
    for (var i in asyncResult.value) {
        bindingString += asyncResult.value[i].id + '\n';
    }
    write('Existing bindings: ' + bindingString);
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="da664-176">La fonction anonyme qui est passée dans la fonction en tant que paramètre est `callback` exécutée lorsque l’opération est terminée.</span><span class="sxs-lookup"><span data-stu-id="da664-176">The anonymous function that is passed into the function as the `callback` parameter is executed when the operation is complete.</span></span> <span data-ttu-id="da664-177">La fonction est appelée avec un seul paramètre, qui contient un `asyncResult` tableau des liaisons dans le document.</span><span class="sxs-lookup"><span data-stu-id="da664-177">The function is called with a single parameter, `asyncResult`, which contains an array of the bindings in the document.</span></span> <span data-ttu-id="da664-178">Le tableau est répété pour générer une chaîne qui contient les ID des liaisons.</span><span class="sxs-lookup"><span data-stu-id="da664-178">The array is iterated to build a string that contains the IDs of the bindings.</span></span> <span data-ttu-id="da664-179">La chaîne est ensuite affichée dans une boîte de message.</span><span class="sxs-lookup"><span data-stu-id="da664-179">The string is then displayed in a message box.</span></span>

## <a name="get-a-binding-by-id-using-the-getbyidasync-method-of-the-bindings-object"></a><span data-ttu-id="da664-180">Obtention d’une liaison par ID en utilisant la méthode getByIdAsync de l’objet Bindings</span><span class="sxs-lookup"><span data-stu-id="da664-180">Get a binding by ID using the getByIdAsync method of the Bindings object</span></span>

<span data-ttu-id="da664-p123">L’exemple suivant indique comment utiliser la méthode [getByIdAsync] pour obtenir une liaison dans un document en spécifiant son ID. Cet exemple suppose qu’une liaison nommée `'myBinding'` a été ajoutée au document à l’aide des méthodes décrites plus haut dans cette rubrique.</span><span class="sxs-lookup"><span data-stu-id="da664-p123">The following example shows how to use the [getByIdAsync] method to get a binding in a document by specifying its ID. This example assumes that a binding named `'myBinding'` was added to the document using one of the methods described earlier in this topic.</span></span>

```js
Office.context.document.bindings.getByIdAsync('myBinding', function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    }
    else {
        write('Retrieved binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="da664-183">Dans l’exemple, le premier `id` paramètre est l’ID de la liaison à récupérer.</span><span class="sxs-lookup"><span data-stu-id="da664-183">In the example, the first `id` parameter is the ID of the binding to retrieve.</span></span>

<span data-ttu-id="da664-184">La fonction anonyme qui est passée dans la fonction en tant que _deuxième_ paramètre de rappel est exécutée lorsque l’opération est terminée.</span><span class="sxs-lookup"><span data-stu-id="da664-184">The anonymous function that is passed into the function as the second _callback_ parameter is executed when the operation is completed.</span></span> <span data-ttu-id="da664-185">La fonction est appelée avec un seul paramètre, _asyncResult_, qui contient le statut de l’appel et la liaison avec l’ID « myBinding ».</span><span class="sxs-lookup"><span data-stu-id="da664-185">The function is called with a single parameter, _asyncResult_, which contains the status of the call and the binding with the ID "myBinding".</span></span>

## <a name="get-a-binding-by-id-using-the-select-method-of-the-office-object"></a><span data-ttu-id="da664-186">Obtention d’une liaison par ID en utilisant la méthode Select de l’objet Office</span><span class="sxs-lookup"><span data-stu-id="da664-186">Get a binding by ID using the select method of the Office object</span></span>

<span data-ttu-id="da664-p125">L’exemple suivant montre comment utiliser la méthode [Office.select] pour obtenir une promesse d’objet [Binding] dans un document en spécifiant son ID dans une chaîne de sélecteur. Il appelle ensuite la méthode [Binding.getDataAsync] pour obtenir des données à partir de la liaison spécifiée. Cet exemple suppose qu’une liaison nommée `'myBinding'` a été ajoutée au document à l’aide des méthodes décrites plus haut dans cette rubrique.</span><span class="sxs-lookup"><span data-stu-id="da664-p125">The following example shows how to use the [Office.select] method to get a [Binding] object promise in a document by specifying its ID in a selector string. It then calls the Binding.[getDataAsync] method to get data from the specified binding. This example assumes that a binding named `'myBinding'` was added to the document using one of the methods described earlier in this topic.</span></span>

```js
Office.select("bindings#myBinding", function onError(){}).getDataAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write(asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


> [!NOTE]
> <span data-ttu-id="da664-190">Si la promesse de méthode renvoie correctement un objet Binding, cet objet expose uniquement les quatre méthodes suivantes de l’objet : `select` [getDataAsync], [setDataAsync], [addHandlerAsync]et [removeHandlerAsync]. []</span><span class="sxs-lookup"><span data-stu-id="da664-190">If the `select` method promise successfully returns a [Binding] object, that object exposes only the following four methods of the object: [getDataAsync], [setDataAsync], [addHandlerAsync], and [removeHandlerAsync].</span></span> <span data-ttu-id="da664-191">Si la promesse ne peut pas renvoyer un objet Binding, le rappel peut être utilisé pour accéder à un objet `onError` .error [asyncResult]pour obtenir plus d’informations. Si vous devez appeler un membre de l’objet Binding autre que les quatre méthodes exposées par la promesse d’objet [Binding] renvoyée par la méthode, utilisez plutôt la méthode `select` [getByIdAsync] à l’aide de la propriété [Document.bindings] et de Bindings.[ méthode getByIdAsync] pour récupérer [l’objet Binding.]</span><span class="sxs-lookup"><span data-stu-id="da664-191">If the promise cannot return a Binding object, the `onError` callback can be used to access an [asyncResult].error object to get more information.If you need to call a member of the Binding object other than the four methods exposed by the [Binding] object promise returned by the `select` method, instead use the [getByIdAsync] method by using the [Document.bindings] property and Bindings.[getByIdAsync] method to retrieve the [Binding] object.</span></span>

## <a name="release-a-binding-by-id"></a><span data-ttu-id="da664-192">Publication d’une liaison par ID</span><span class="sxs-lookup"><span data-stu-id="da664-192">Release a binding by ID</span></span>

<span data-ttu-id="da664-193">L’exemple suivant montre comment utiliser la méthode [releaseByIdAsync] pour publier une liaison dans un document en spécifiant son ID.</span><span class="sxs-lookup"><span data-stu-id="da664-193">The following example shows how use the [releaseByIdAsync] method to release a binding in a document by specifying its ID.</span></span>

```js
Office.context.document.bindings.releaseByIdAsync('myBinding', function (asyncResult) {
    write('Released myBinding!');
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="da664-194">Dans l’exemple, le premier paramètre `id` est l’ID de la liaison à publier.</span><span class="sxs-lookup"><span data-stu-id="da664-194">In the example, the first `id` parameter is the ID of the binding to release.</span></span>

<span data-ttu-id="da664-p127">La fonction anonyme qui est passée dans la fonction comme le deuxième paramètre est un rappel qui est exécuté lorsque l’opération est terminée. La fonction est appelée avec un seul paramètre,  [asyncResult], qui contient le statut de l’appel.</span><span class="sxs-lookup"><span data-stu-id="da664-p127">The anonymous function that is passed into the function as the second parameter is a callback that is executed when the operation is complete. The function is called with a single parameter,  [asyncResult], which contains the status of the call.</span></span>

## <a name="read-data-from-a-binding"></a><span data-ttu-id="da664-197">Lecture de données à partir d’une liaison</span><span class="sxs-lookup"><span data-stu-id="da664-197">Read data from a binding</span></span>

<span data-ttu-id="da664-198">L’exemple suivant montre comment utiliser la méthode [getDataAsync] pour obtenir des données à partir d’une liaison existante.</span><span class="sxs-lookup"><span data-stu-id="da664-198">The following example shows how to use the [getDataAsync] method to get data from an existing binding.</span></span>

```js
myBinding.getDataAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write(asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="da664-p128">`myBinding` est une variable qui contient une liaison de texte existante dans le document. Vous pouvez également utiliser [Office.select] pour accéder à la liaison avec son identifiant et commencer à appeler la méthode [getDataAsync] de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="da664-p128">`myBinding` is a variable that contains an existing text binding in the document. Alternatively, you could use the [Office.select] to access the binding by its ID, and start your call to the [getDataAsync] method, like this:</span></span> 

```js
Office.select("bindings#myBindingID").getDataAsync
```

<span data-ttu-id="da664-p129">La fonction anonyme qui est passée dans la fonction est un rappel qui est exécuté lorsque l’opération est terminée. La propriété [AsyncResult].value contient les données dans `myBinding`. Le type de valeur dépend du type de liaison. La liaison dans cet exemple est une liaison de texte. Par conséquent, la valeur contiendra une chaîne. Pour obtenir des exemples supplémentaires concernant l’utilisation des liaisons de matrice et de tableau, consultez la rubrique sur la méthode [getDataAsync].</span><span class="sxs-lookup"><span data-stu-id="da664-p129">The anonymous function that is passed into the function is a callback that is executed when the operation is complete. The [AsyncResult].value property contains the data within `myBinding`. The type of the value depends on the binding type. The binding in this example is a text binding. Therefore, the value will contain a string. For additional examples of working with matrix and table bindings, see the [getDataAsync] method topic.</span></span>

## <a name="write-data-to-a-binding"></a><span data-ttu-id="da664-207">Écriture de données dans une liaison</span><span class="sxs-lookup"><span data-stu-id="da664-207">Write data to a binding</span></span>

<span data-ttu-id="da664-208">L’exemple suivant montre comment utiliser la méthode [setDataAsync] pour définir des données dans une liaison existante.</span><span class="sxs-lookup"><span data-stu-id="da664-208">The following example shows how to use the [setDataAsync] method to set data in an existing binding.</span></span>

```js
myBinding.setDataAsync('Hello World!', function (asyncResult) { });
```

<span data-ttu-id="da664-209">`myBinding` est une variable qui contient une liaison de texte existante dans le document.</span><span class="sxs-lookup"><span data-stu-id="da664-209">`myBinding` is a variable that contains an existing text binding in the document.</span></span>

<span data-ttu-id="da664-210">Dans l’exemple, le premier paramètre est la valeur à définir `myBinding` sur .</span><span class="sxs-lookup"><span data-stu-id="da664-210">In the example, the first parameter is the value to set on `myBinding`.</span></span> <span data-ttu-id="da664-211">Comme il s’agit d’une liaison de texte, la valeur est de type `string`.</span><span class="sxs-lookup"><span data-stu-id="da664-211">Because this is a text binding, the value is a `string`.</span></span> <span data-ttu-id="da664-212">Différents types de liaisons acceptent divers types de données.</span><span class="sxs-lookup"><span data-stu-id="da664-212">Different binding types accept different types of data.</span></span>

<span data-ttu-id="da664-213">La fonction anonyme qui est passée dans la fonction est un rappel qui est exécuté lorsque l’opération est terminée.</span><span class="sxs-lookup"><span data-stu-id="da664-213">The anonymous function that is passed into the function is a callback that is executed when the operation is complete.</span></span> <span data-ttu-id="da664-214">La fonction est appelée avec un seul paramètre, `asyncResult` qui contient l’état du résultat.</span><span class="sxs-lookup"><span data-stu-id="da664-214">The function is called with a single parameter, `asyncResult`, which contains the status of the result.</span></span>

> [!NOTE]
> <span data-ttu-id="da664-215">Depuis la publication d’Excel 2013 SP1 et de la version correspondante d’Excel sur le web, vous pouvez désormais [définir la mise en forme lors de l’écriture et de la mise à jour des données dans des tableaux liés](../excel/excel-add-ins-tables.md).</span><span class="sxs-lookup"><span data-stu-id="da664-215">Starting with the release of the Excel 2013 SP1 and the corresponding build of Excel on the web, you can now [set formatting when writing and updating data in bound tables](../excel/excel-add-ins-tables.md).</span></span>

## <a name="detect-changes-to-data-or-the-selection-in-a-binding"></a><span data-ttu-id="da664-216">Détection des modifications apportées aux données ou à la section dans une liaison</span><span class="sxs-lookup"><span data-stu-id="da664-216">Detect changes to data or the selection in a binding</span></span>

<span data-ttu-id="da664-217">L’exemple suivant montre comment lier un gestionnaire d’événements à l’événement [DataChanged](/javascript/api/office/office.binding) d’une liaison ayant l’ID « MyBinding ».</span><span class="sxs-lookup"><span data-stu-id="da664-217">The following example shows how to attach an event handler to the [DataChanged](/javascript/api/office/office.binding) event of a binding with an id of "MyBinding".</span></span>

```js
function addHandler() {
Office.select("bindings#MyBinding").addHandlerAsync(
    Office.EventType.BindingDataChanged, dataChanged);
}
function dataChanged(eventArgs) {
    write('Bound data changed in binding: ' + eventArgs.binding.id);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

<span data-ttu-id="da664-218">`myBinding` est une variable qui contient une liaison de texte existante dans le document.</span><span class="sxs-lookup"><span data-stu-id="da664-218">The `myBinding` is a variable that contains an existing text binding in the document.</span></span>

<span data-ttu-id="da664-219">Le premier _paramètre eventType_ de la méthode [addHandlerAsync] spécifie le nom de l’événement à abonner.</span><span class="sxs-lookup"><span data-stu-id="da664-219">The first _eventType_ parameter of the [addHandlerAsync] method specifies the name of the event to subscribe to.</span></span> <span data-ttu-id="da664-220">[Office.EventType] est une énumération des valeurs de types d’événement disponibles.</span><span class="sxs-lookup"><span data-stu-id="da664-220">[Office.EventType] is an enumeration of available event type values.</span></span> <span data-ttu-id="da664-221">`Office.EventType.BindingDataChanged` évalue la chaîne « bindingDataChanged ».</span><span class="sxs-lookup"><span data-stu-id="da664-221">`Office.EventType.BindingDataChanged` evaluates to the string "bindingDataChanged".</span></span>

<span data-ttu-id="da664-222">La fonction qui est passée dans la fonction en tant que deuxième paramètre de handler est un handler d’événements qui est exécuté lorsque les données de la `dataChanged` liaison sont  modifiées.</span><span class="sxs-lookup"><span data-stu-id="da664-222">The `dataChanged` function that is passed into the function as the second _handler_ parameter is an event handler that is executed when the data in the binding is changed.</span></span> <span data-ttu-id="da664-223">La fonction est appelée avec un seul paramètre, _eventArgs_, qui contient une référence à la liaison.</span><span class="sxs-lookup"><span data-stu-id="da664-223">The function is called with a single parameter, _eventArgs_, which contains a reference to the binding.</span></span> <span data-ttu-id="da664-224">Cette liaison peut être utilisée pour récupérer les données mises à jour.</span><span class="sxs-lookup"><span data-stu-id="da664-224">This binding can be used to retrieve the updated data.</span></span>

<span data-ttu-id="da664-p134">De même, vous pouvez détecter lorsqu’un utilisateur modifie la sélection dans une liaison en ajoutant un gestionnaire d’événements à l’événement [SelectionChanged] d’une liaison. Pour ce faire, spécifiez le paramètre `eventType` de la méthode [addHandlerAsync] comme `Office.EventType.BindingSelectionChanged` ou `"bindingSelectionChanged"`.</span><span class="sxs-lookup"><span data-stu-id="da664-p134">Similarly, you can detect when a user changes selection in a binding by attaching an event handler to the [SelectionChanged] event of a binding. To do that, specify the `eventType` parameter of the [addHandlerAsync] method as `Office.EventType.BindingSelectionChanged` or `"bindingSelectionChanged"`.</span></span>

<span data-ttu-id="da664-p135">Vous pouvez ajouter plusieurs gestionnaires d’événements pour un événement donné en appelant à nouveau la méthode [addHandlerAsync] et en transmettant une fonction de gestionnaire d’événements supplémentaire pour le paramètre `handler`. Cela fonctionnera correctement tant que le nom de chaque fonction de gestionnaire d’événements est unique.</span><span class="sxs-lookup"><span data-stu-id="da664-p135">You can add multiple event handlers for a given event by calling the [addHandlerAsync] method again and passing in an additional event handler function for the `handler` parameter. This will work correctly as long as the name of each event handler function is unique.</span></span>

### <a name="remove-an-event-handler"></a><span data-ttu-id="da664-229">Suppression d’un gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="da664-229">Remove an event handler</span></span>

<span data-ttu-id="da664-p136">Pour supprimer un gestionnaire d’événements pour un événement, appelez la méthode [removeHandlerAsync] en transmettant le type d’événement en tant que premier paramètre _eventType_, puis le nom de la fonction de gestionnaire d’événements à supprimer comme deuxième paramètre _handler_. Par exemple, la fonction suivante supprimera la fonction de gestionnaire d’événements `dataChanged` ajoutée dans l’exemple de la section précédente.</span><span class="sxs-lookup"><span data-stu-id="da664-p136">To remove an event handler for an event, call the [removeHandlerAsync] method passing in the event type as the first _eventType_ parameter, and the name of the event handler function to remove as the second _handler_ parameter. For example, the following function will remove the `dataChanged` event handler function added in the previous section's example.</span></span>

```js
function removeEventHandlerFromBinding() {
    Office.select("bindings#MyBinding").removeHandlerAsync(
        Office.EventType.BindingDataChanged, {handler:dataChanged});
}
```

> [!IMPORTANT]
> <span data-ttu-id="da664-232">Si le paramètre de _handler_ facultatif est omis lors de l’appel de la méthode [removeHandlerAsync,] tous les handlers d’événements pour le paramètre spécifié `eventType` sont supprimés.</span><span class="sxs-lookup"><span data-stu-id="da664-232">If the optional _handler_ parameter is omitted when the [removeHandlerAsync] method is called, all event handlers for the specified `eventType` will be removed.</span></span>

## <a name="see-also"></a><span data-ttu-id="da664-233">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="da664-233">See also</span></span>

- [<span data-ttu-id="da664-234">Compréhension de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="da664-234">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="da664-235">Programmation asynchrone dans des compléments Office</span><span class="sxs-lookup"><span data-stu-id="da664-235">Asynchronous programming in Office Add-ins</span></span>](asynchronous-programming-in-office-add-ins.md)
- [<span data-ttu-id="da664-236">Lecture et écriture de données dans la sélection active d’un document ou d’une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="da664-236">Read and write data to the active selection in a document or spreadsheet</span></span>](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)

[Binding]:               /javascript/api/office/office.binding
[MatrixBinding]:         /javascript/api/office/office.matrixbinding
[TableBinding]:          /javascript/api/office/office.tablebinding
[TextBinding]:           /javascript/api/office/office.textbinding
[getDataAsync]:          /javascript/api/office/Office.Binding#getdataasync-options--callback-
[setDataAsync]:          /javascript/api/office/Office.Binding#setdataasync-data--options--callback-
[SelectionChanged]:      /javascript/api/office/office.bindingselectionchangedeventargs
[addHandlerAsync]:       /javascript/api/office/Office.Binding#addhandlerasync-eventtype--handler--options--callback-
[removeHandlerAsync]:    /javascript/api/office/Office.Binding#removehandlerasync-eventtype--options--callback-

[Bindings]:              /javascript/api/office/office.bindings
[getByIdAsync]:          /javascript/api/office/office.bindings#getbyidasync-id--options--callback- 
[getAllAsync]:           /javascript/api/office/office.bindings#getallasync-options--callback-
[addFromNamedItemAsync]: /javascript/api/office/office.bindings#addfromnameditemasync-itemname--bindingtype--options--callback-
[addFromSelectionAsync]: /javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-
[addFromPromptAsync]:    /javascript/api/office/office.bindings#addfrompromptasync-bindingtype--options--callback-
[releaseByIdAsync]:      /javascript/api/office/office.bindings#releasebyidasync-id--options--callback-

[AsyncResult]:          /javascript/api/office/office.asyncresult
[Office.BindingType]:   /javascript/api/office/office.bindingtype
[Office.select]:        /javascript/api/office 
[Office.EventType]:     /javascript/api/office/office.eventtype 
[Document.bindings]:    /javascript/api/office/office.document

[TableBinding.rowCount]: /javascript/api/office/office.tablebinding
[BindingSelectionChangedEventArgs]: /javascript/api/office/office.bindingselectionchangedeventargs
