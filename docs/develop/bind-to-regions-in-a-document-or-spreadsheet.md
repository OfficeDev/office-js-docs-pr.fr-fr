---
title: Lier des r?gions dans un document ou une feuille de calcul
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: bd26aa12e5d6da145fb6a2a89daf937cf6e88f04
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="bind-to-regions-in-a-document-or-spreadsheet"></a><span data-ttu-id="2c0f1-102">Lier des r?gions dans un document ou une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="2c0f1-102">Bind to regions in a document or spreadsheet</span></span>

<span data-ttu-id="2c0f1-p101">L?acc?s aux donn?es bas?es sur une liaison permet aux compl?ments de contenu et du volet Office d?acc?der de fa?on coh?rente ? une zone particuli?re d?un document ou d?une feuille de calcul au moyen d?un identificateur. Le compl?ment doit d?abord ?tablir la liaison en appelant l?une des m?thodes qui associent une partie du document ? un identificateur unique : [addFromPromptAsync], [addFromSelectionAsync] ou [addFromNamedItemAsync]. Une fois la liaison ?tablie, le compl?ment peut utiliser l?identificateur fourni pour acc?der aux donn?es contenues dans la zone associ?e du document ou de la feuille de calcul. La cr?ation de liaisons apporte la valeur ajout?e suivante ? votre compl?ment :</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p101">Binding-based data access enables content and task pane add-ins to consistently access a particular region of a document or spreadsheet through an identifier. The add-in first needs to establish the binding by calling one of the methods that associates a portion of the document with a unique identifier: [addFromPromptAsync], [addFromSelectionAsync], or [addFromNamedItemAsync]. After the binding is established, the add-in can use the provided identifier to access the data contained in the associated region of the document or spreadsheet. Creating bindings provides the following value to your add-in:</span></span>


- <span data-ttu-id="2c0f1-107">Elle permet l?acc?s aux structures de donn?es communes sur les applications Office prises en charge, telles que : tableaux, plages ou texte (s?rie contigu? de caract?res).</span><span class="sxs-lookup"><span data-stu-id="2c0f1-107">Permits access to common data structures across supported Office applications, such as: tables, ranges, or text (a contiguous run of characters).</span></span>
    
- <span data-ttu-id="2c0f1-108">Elle permet les op?rations de lecture/?criture sans exiger que l?utilisateur effectue une s?lection.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-108">Enables read/write operations without requiring the user to make a selection.</span></span>
    
- <span data-ttu-id="2c0f1-p102">Elle ?tablit une relation entre le compl?ment et les donn?es du document. Les liaisons persistent dans le document et sont accessibles par la suite.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p102">Establishes a relationship between the add-in and the data in the document. Bindings are persisted in the document, and can be accessed at a later time.</span></span>
    
<span data-ttu-id="2c0f1-p103">L??tablissement d?une liaison vous permet ?galement de vous abonner aux donn?es et aux ?v?nements de changement de s?lection qui sont concern?s par cette r?gion particuli?re du document ou de la feuille de calcul. Cela signifie que le compl?ment est seulement notifi? des changements qui surviennent dans la r?gion d?limit?e, par opposition aux changements g?n?raux affectant l?ensemble du document ou de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p103">Establishing a binding also allows you to subscribe to data and selection change events that are scoped to that particular region of the document or spreadsheet. This means that the add-in is only notified of changes that happen within the bound region as opposed to general changes across the whole document or spreadsheet.</span></span>

<span data-ttu-id="2c0f1-p104">L?objet [Bindings] expose une m?thode [getAllAsync] qui donne acc?s ? toutes les liaisons ?tablies dans le document ou la feuille de calcul. Une liaison individuelle est accessible par son ID ? l?aide de la m?thode [Bindings.getBindingByIdAsync] ou [Office.select]. Vous pouvez ?tablir de nouvelles liaisons et supprimer des liaisons existantes en utilisant l?une des m?thodes suivantes de l?objet [Bindings] : [addFromSelectionAsync], [addFromPromptAsync], [addFromNamedItemAsync] ou [releaseByIdAsync].</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p104">The [Bindings] object exposes a [getAllAsync] method that gives access to the set of all bindings established on the document or spreadsheet. An individual binding can be accessed by its ID using either the Bindings.[getByIdAsync] or [Office.select] methods. You can establish new bindings as well as remove existing ones by using one of the following methods of the [Bindings] object: [addFromSelectionAsync], [addFromPromptAsync], [addFromNamedItemAsync], or [releaseByIdAsync].</span></span>


## <a name="binding-types"></a><span data-ttu-id="2c0f1-116">Types de liaison</span><span class="sxs-lookup"><span data-stu-id="2c0f1-116">Binding types</span></span>

<span data-ttu-id="2c0f1-117">Vous sp?cifiez [trois types de liaisons diff?rents][Office.BindingType] avec le param?tre _bindingType_ lorsque vous cr?ez une liaison avec les m?thodes [addFromSelectionAsync], [addFromPromptAsync] ou [addFromNamedItemAsync] :</span><span class="sxs-lookup"><span data-stu-id="2c0f1-117">There are [three different types of bindings][Office.BindingType] that you specify with the  _bindingType_ parameter when you create a binding with the [addFromSelectionAsync], [addFromPromptAsync] or [addFromNamedItemAsync] methods:</span></span>

1. <span data-ttu-id="2c0f1-118">**[Liaison de texte][TextBinding]** - ?tablit une liaison ? une zone du document qui est repr?sent?e en tant que texte.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-118">**[Text Binding][TextBinding]** - Binds to a region of the document that can be represented as text.</span></span>

    <span data-ttu-id="2c0f1-p105">Dans Word, la plupart des s?lections contigu?s sont valides, tandis que dans Excel, seules les s?lections de cellules uniques peuvent ?tre la cible d?une liaison de texte. Dans Excel, seul le texte brut est pris en charge. Dans Word, trois formats sont pris en charge : texte brut, HTML et Open XML pour Office.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p105">In Word, most contiguous selections are valid, while in Excel only single cell selections can be the target of a text binding. In Excel, only plain text is supported. In Word, three formats are supported: plain text, HTML, and Open XML for Office.</span></span>

2. <span data-ttu-id="2c0f1-p106">**[Liaison de matrice][MatrixBinding]** - ?tablit une liaison ? une zone d?un document qui contient des donn?es tabulaires sans en-t?te. Les donn?es dans une liaison de matrice sont ?crites ou lues comme un **tableau** bidimensionnel, ce qui est impl?ment? sous la forme d?un tableau de tableaux dans JavaScript. Par exemple, deux lignes d?une valeur de **cha?ne** dans deux colonnes peuvent ?tre ?crites ou lues comme ` [['a', 'b'], ['c', 'd']]`, et une colonne unique de trois lignes peut ?tre ?crite ou lue comme `[['a'], ['b'], ['c']]`.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p106">**[Matrix Binding][MatrixBinding]** - Binds to a fixed region of a document that contains tabular data without headers.Data in a matrix binding is written or read as a two dimensional  **Array**, which in JavaScript is implemented as an array of arrays. For example, two rows of  **string** values in two columns can be written or read as ` [['a', 'b'], ['c', 'd']]`, and a single column of three rows can be written or read as  `[['a'], ['b'], ['c']]`.</span></span>

    <span data-ttu-id="2c0f1-p107">Dans Excel, toute s?lection contigu? de cellules peut ?tre utilis?e pour ?tablir une liaison de matrice. Dans Word, seuls les tableaux prennent en charge la liaison de matrice.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p107">In Excel, any contiguous selection of cells can be used to establish a matrix binding. In Word, only tables support matrix binding.</span></span>

3. <span data-ttu-id="2c0f1-p108">**[Liaison de tableau][TableBinding]** - ?tablit une liaison ? une zone d?un document qui contient un tableau avec des en-t?tes. Les donn?es dans une liaison de tableau sont ?crites ou lues comme un objet [TableData](https://dev.office.com/reference/add-ins/shared/tabledata). L?objet `TableData` expose les donn?es via les propri?t?s `headers` et `rows`.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p108">**[Table Binding][TableBinding]** - Binds to a region of a document that contains a table with headers.Data in a table binding is written or read as a [TableData](https://dev.office.com/reference/add-ins/shared/tabledata) object. The `TableData` object exposes the data through the `headers` and `rows` properties.</span></span>

    <span data-ttu-id="2c0f1-p109">Tout tableau Excel ou Word peut ?tre la base d?une liaison de tableau. Une fois que vous ?tablissez une liaison de tableau, chaque nouvelle ligne ou colonne qu?un utilisateur ajoute au tableau est automatiquement incluse dans la liaison.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p109">Any Excel or Word table can be the basis for a table binding. After you establish a table binding, each new row or column a user adds to the table is automatically included in the binding.</span></span>

<span data-ttu-id="2c0f1-p110">Apr?s la cr?ation d?une liaison ? l?aide de l?une des trois m?thodes ? addFrom ? de l?objet `Bindings`, vous pouvez travailler avec les donn?es et les propri?t?s de la liaison en utilisant les m?thodes de l?objet correspondant : [MatrixBinding], [TableBinding] ou [TextBinding]. Ces trois objets h?ritent des m?thodes [getDataAsync] et [setDataAsync] de l?objet `Binding` qui vous permettent d?interagir avec les donn?es li?es.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p110">After a binding is created by using one of the three "addFrom" methods of the  `Bindings` object, you can work with the binding's data and properties by using the methods of the corresponding object: [MatrixBinding], [TableBinding], or [TextBinding]. All three of these objects inherit the [getDataAsync] and [setDataAsync] methods of the `Binding` object that enable you to interact with the bound data.</span></span>

> [!NOTE]
> <span data-ttu-id="2c0f1-p111">**Quand devez-vous utiliser une liaison de matrice ou une liaison de tableau ?** Lorsque les donn?es tabulaires avec lesquelles vous travaillez contiennent une ligne de total, vous devez utiliser une liaison de matrice si le script de votre compl?ment doit acc?der aux valeurs figurant dans la ligne de total ou d?tecter que la s?lection de l?utilisateur figure dans la ligne de total. Si vous ?tablissez une liaison de tableau pour des donn?es tabulaires qui contiennent une ligne de total, la propri?t? [TableBinding.rowCount] et les propri?t?s `rowCount` et `startRow` de l?objet [BindingSelectionChangedEventArgs] dans les gestionnaires d??v?nements ne refl?teront pas la ligne de total dans leurs valeurs. Pour contourner cette limitation, vous devez ?tablir une liaison de matrice pour travailler avec la ligne de total.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p111">**When should you use matrix versus table bindings?** When the tabular data you are working with contains a total row, you must use a matrix binding if your add-in's script needs to access values in the total row or detect that the user's selection is in the total row. If you establish a table binding for tabular data that contains a total row, the [TableBinding.rowCount] property and the `rowCount` and `startRow` properties of the [BindingSelectionChangedEventArgs] object in event handlers won't reflect the total row in their values. To work around this limitation, you must use establish a matrix binding to work with the total row.</span></span>

## <a name="add-a-binding-to-the-users-current-selection"></a><span data-ttu-id="2c0f1-136">Ajout d?une liaison ? la s?lection actuelle de l?utilisateur</span><span class="sxs-lookup"><span data-stu-id="2c0f1-136">Add a binding to the user's current selection</span></span>

<span data-ttu-id="2c0f1-137">L?exemple suivant montre comment ajouter une liaison de texte nomm?e `myBinding` ? la s?lection actuelle dans un document ? l?aide de la m?thode [addFromSelectionAsync].</span><span class="sxs-lookup"><span data-stu-id="2c0f1-137">The following example shows how to add a text binding called  `myBinding` to the current selection in a document by using the [addFromSelectionAsync] method.</span></span>


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

<span data-ttu-id="2c0f1-p112">Dans cet exemple, le type de liaison sp?cifi? est ? Text ?. Cela signifie qu?un objet [TextBinding] sera cr?? pour la s?lection. Diff?rents types de liaison exposent diff?rentes donn?es et op?rations. [Office.BindingType] est une ?num?ration des valeurs de types de liaison disponibles.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p112">In this example, the specified binding type is text. This means that a [TextBinding] will be created for the selection. Different binding types expose different data and operations. [Office.BindingType] is an enumeration of available binding type values.</span></span>

<span data-ttu-id="2c0f1-p113">Le deuxi?me param?tre facultatif est un objet qui sp?cifie l?ID de la nouvelle liaison cr??e. Si un ID n?est pas sp?cifi?, un ID est g?n?r? automatiquement.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p113">The second optional parameter is an object that specifies the ID of the new binding being created. If an ID is not specified, one is generated automatically.</span></span>

<span data-ttu-id="2c0f1-p114">La fonction anonyme qui est pass?e dans la fonction comme param?tre final _callback_ est ex?cut?e lorsque la cr?ation de la liaison est termin?e. La fonction est appel?e avec un seul param?tre, `asyncResult`, ce qui donne acc?s ? un objet [AsyncResult] qui fournit l??tat de l?appel. La propri?t? `AsyncResult.value` contient une r?f?rence ? un objet [Binding] du type sp?cifi? pour la liaison cr??e r?cemment. Vous pouvez utiliser cet objet [Binding] pour obtenir et d?finir les donn?es.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p114">The anonymous function that is passed into the function as the final  _callback_ parameter is executed when the creation of the binding is complete. The function is called with a single parameter, `asyncResult`, which provides access to an [AsyncResult] object that provides the status of the call. The `AsyncResult.value` property contains a reference to a [Binding] object of the type that is specified for the newly created binding. You can use this [Binding] object to get and set data.</span></span>

## <a name="add-a-binding-from-a-prompt"></a><span data-ttu-id="2c0f1-148">Ajout d?une liaison ? partir d?une invite</span><span class="sxs-lookup"><span data-stu-id="2c0f1-148">Add a binding from a prompt</span></span>

<span data-ttu-id="2c0f1-p115">L?exemple suivant indique comment ajouter une liaison de texte appel?e `myBinding` ? l?aide de la m?thode [addFromPromptAsync]. Cette m?thode permet ? l?utilisateur de sp?cifier la plage pour la liaison ? l?aide de l?invite de s?lection de plage int?gr?e.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p115">The following example shows how to add a text binding called  `myBinding` by using the [addFromPromptAsync] method. This method lets the user specify the range for the binding by using the application's built-in range selection prompt.</span></span>


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

<span data-ttu-id="2c0f1-p116">Dans cet exemple, le type de liaison sp?cifi? est ? Text ?. Cela signifie qu?un objet [TextBinding] sera cr?? pour la s?lection que l?utilisateur sp?cifie dans l?invite.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p116">In this example, the specified binding type is text. This means that a [TextBinding] will be created for the selection that the user specifies in the prompt.</span></span>

<span data-ttu-id="2c0f1-p117">Le deuxi?me param?tre est un objet qui contient l?ID de la nouvelle liaison cr??e. Si un ID n?est pas sp?cifi?, un ID est g?n?r? automatiquement.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p117">The second parameter is an object that contains the ID of the new binding being created. If an ID is not specified, one is generated automatically.</span></span>

<span data-ttu-id="2c0f1-p118">La fonction anonyme transmise dans la fonction comme troisi?me param?tre _callback_ est ex?cut?e lorsque la cr?ation de la liaison est termin?e. Lorsque la fonction de rappel s?ex?cute, l?objet [AsyncResult] contient le statut de l?appel et la nouvelle liaison.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p118">The anonymous function passed into the function as the third  _callback_ parameter is executed when the creation of the binding is complete. When the callback function executes, the [AsyncResult] object contains the status of the call and the newly created binding.</span></span>

<span data-ttu-id="2c0f1-157">La figure 1 montre l?invite de s?lection de plage int?gr?e dans Excel.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-157">Figure 1 shows the built-in range selection prompt in Excel.</span></span>


<span data-ttu-id="2c0f1-158">*Figure 1. Interface utilisateur de s?lection de donn?es dans Excel*</span><span class="sxs-lookup"><span data-stu-id="2c0f1-158">*Figure 1. Excel Select Data UI*</span></span>

![Interface utilisateur de s?lection de donn?es dans Excel](../images/agave-api-overview-excel-selection-ui.png)


## <a name="add-a-binding-to-a-named-item"></a><span data-ttu-id="2c0f1-160">Ajout d?une liaison ? un ?l?ment nomm?</span><span class="sxs-lookup"><span data-stu-id="2c0f1-160">Add a binding to a named item</span></span>


<span data-ttu-id="2c0f1-161">L?exemple suivant montre comment ajouter une liaison de matrice ? l??l?ment nomm? `myRange` existant en utilisant la m?thode [addFromNamedItemAsync], et d?finit le param?tre `id` de la liaison sur ? myMatrix ?.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-161">The following example shows how to add a binding to the existing  `myRange` named item as a "matrix" binding by using the [addFromNamedItemAsync] method, and assigns the binding's `id` as "myMatrix".</span></span>


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

<span data-ttu-id="2c0f1-p119">**Pour Excel**, le param?tre `itemName` de la m?thode [addFromNamedItemAsync] peut se r?f?rer ? une plage nomm?e existante, une plage sp?cifi?e avec le style de r?f?rence `A1` `("A1:A3")` ou un tableau. Par d?faut, l?ajout d?un tableau dans Excel entra?ne l?affectation du nom ? Tableau1 ? pour le premier tableau que vous ajoutez, ? Tableau2 ? pour le deuxi?me tableau que vous ajoutez, et ainsi de suite. Pour affecter un nom significatif ? un tableau dans l?interface utilisateur d?Excel, servez-vous de la propri?t? **Table Name** sous l?onglet **Outils de tableau | Conception** du ruban.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p119">**For Excel**, the  `itemName` parameter of the [addFromNamedItemAsync] method can refer to an existing named range, a range specified with the `A1` reference style `("A1:A3")`, or a table. By default, adding a table in Excel assigns the name "Table1" for the first table you add, "Table2" for the second table you add, and so on. To assign a meaningful name for a table in the Excel UI, use the **Table Name** property on the **Table Tools | Design** tab of the ribbon.</span></span>


> [!NOTE]
> <span data-ttu-id="2c0f1-165">Dans Excel, lors de la sp?cification d?un tableau comme ?l?ment nomm?, vous devez enti?rement qualifier le nom pour inclure le nom de la feuille de calcul dans le nom du tableau dans ce format :  `"Sheet1!Table1"`</span><span class="sxs-lookup"><span data-stu-id="2c0f1-165">In Excel, when specifying a table as a named item, you must fully qualify the name to include the worksheet name in the name of the table in this format:  `"Sheet1!Table1"`</span></span>

<span data-ttu-id="2c0f1-166">L?exemple suivant cr?e une liaison dans Excel aux trois premi?res cellules de la colonne A (`"A1:A3"`), attribue l?id`"MyCities"`, puis ?crit trois noms de ville dans cette liaison.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-166">The following example creates a binding in Excel to the first three cells in column A ( `"A1:A3"`), assigns the  id `"MyCities"`, and then writes three city names to that binding.</span></span>


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

<span data-ttu-id="2c0f1-p120">**Pour Word**, le param?tre `itemName` de la m?thode [addFromNamedItemAsync] fait r?f?rence ? la propri?t? `Title` d?un contr?le de contenu `Rich Text`. (Vous ne pouvez r?aliser de liaison avec des contr?les de contenu diff?rents du contr?le de contenu `Rich Text`.)</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p120">**For Word**, the  `itemName` parameter of the [addFromNamedItemAsync] method refers to the `Title` property of a `Rich Text` content control. (You can't bind to content controls other than the `Rich Text` content control.)</span></span>

<span data-ttu-id="2c0f1-p121">Par d?faut, un contr?le de contenu ne comporte aucune valeur affect?e `Title*`. Pour attribuer un nom significatif dans l?interface utilisateur de Word, apr?s avoir ins?r? un contr?le de contenu de **texte enrichi** ? partir du groupe **Contr?les** sous l?onglet **D?veloppeur** du ruban, utilisez la commande **Propri?t?s** dans le groupe **Contr?les** pour afficher la bo?te de dialogue **Propri?t?s du contr?le de contenu**. D?finissez la propri?t? **Title** du contr?le de contenu sur le nom auquel vous souhaitez faire r?f?rence ? partir de votre code.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p121">By default, a content control has no  `Title*`value assigned. To assign a meaningful name in the Word UI, after inserting a **Rich Text** content control from the **Controls** group on the **Developer** tab of the ribbon, use the **Properties** command in the **Controls** group to display the **Content Control Properties** dialog box. Then set the **Title** property of the content control to the name you want to reference from your code.</span></span>

<span data-ttu-id="2c0f1-172">L?exemple suivant cr?e une liaison de texte dans Word vers un contr?le de contenu de texte enrichi nomm?  `"FirstName"`, attribue l? **id**`"firstName"`, puis affiche cette information.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-172">The following example creates a text binding in Word to a rich text content control named  `"FirstName"`, assigns the  **id** `"firstName"`, and then displays that information.</span></span>


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

## <a name="get-all-bindings"></a><span data-ttu-id="2c0f1-173">Obtention de toutes les liaisons</span><span class="sxs-lookup"><span data-stu-id="2c0f1-173">Get all bindings</span></span>


<span data-ttu-id="2c0f1-174">L?exemple suivant montre comment obtenir toutes les liaisons dans un document en utilisant la m?thode Bindings.[getAllAsync].</span><span class="sxs-lookup"><span data-stu-id="2c0f1-174">The following example shows how to get all bindings in a document by using the Bindings.[getAllAsync] method.</span></span>


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

<span data-ttu-id="2c0f1-p122">La fonction anonyme qui est pass?e dans la fonction comme param?tre `callback` est ex?cut?e lorsque l?op?ration est termin?e. La fonction est appel?e avec un seul param?tre, `asyncResult`, qui contient un tableau des liaisons dans le document. Le tableau est r?p?t? pour g?n?rer une cha?ne qui contient les ID des liaisons. La cha?ne est ensuite affich?e dans une bo?te de message.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p122">The anonymous function that is passed into the function as the  `callback` parameter is executed when the operation is complete. The function is called with a single parameter, `asyncResult`, which contains an  array of the bindings in the document. The array is iterated to build a string that contains the IDs of the bindings. The string is then displayed in a message box.</span></span>


## <a name="get-a-binding-by-id-using-the-getbyidasync-method-of-the-bindings-object"></a><span data-ttu-id="2c0f1-179">Obtention d?une liaison par ID en utilisant la m?thode getByIdAsync de l?objet Bindings</span><span class="sxs-lookup"><span data-stu-id="2c0f1-179">Get a binding by ID using the getByIdAsync method of the Bindings object</span></span>


<span data-ttu-id="2c0f1-p123">L?exemple suivant indique comment utiliser la m?thode [getByIdAsync] pour obtenir une liaison dans un document en sp?cifiant son ID. Cet exemple suppose qu?une liaison nomm?e `'myBinding'` a ?t? ajout?e au document ? l?aide des m?thodes d?crites plus haut dans cette rubrique.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p123">The following example shows how to use the [getByIdAsync] method to get a binding in a document by specifying its ID. This example assumes that a binding named `'myBinding'` was added to the document using one of the methods described earlier in this topic.</span></span>


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

<span data-ttu-id="2c0f1-182">Dans l?exemple, le premier param?tre `id` est l?ID de la liaison ? r?cup?rer.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-182">In the example, the first  `id` parameter is the ID of the binding to retrieve.</span></span>

<span data-ttu-id="2c0f1-p124">La fonction anonyme qui est pass?e dans la fonction comme second param?tre  _callback_ est ex?cut?e lorsque l?op?ration est termin?e. La fonction est appel?e avec un seul param?tre, _asyncResult_, qui contient le statut de l?appel et la liaison avec l?ID ? myBinding ?.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p124">The anonymous function that is passed into the function as the second  _callback_ parameter is executed when the operation is completed. The function is called with a single parameter, _asyncResult_, which contains the status of the call and the binding with the ID "myBinding".</span></span>


## <a name="get-a-binding-by-id-using-the-select-method-of-the-office-object"></a><span data-ttu-id="2c0f1-185">Obtention d?une liaison par ID en utilisant la m?thode Select de l?objet Office</span><span class="sxs-lookup"><span data-stu-id="2c0f1-185">Get a binding by ID using the select method of the Office object</span></span>


<span data-ttu-id="2c0f1-p125">L?exemple suivant montre comment utiliser la m?thode [Office.select] pour obtenir une promesse d?objet [Binding] dans un document en sp?cifiant son ID dans une cha?ne de s?lecteur. Il appelle ensuite la m?thode [Binding.getDataAsync] pour obtenir des donn?es ? partir de la liaison sp?cifi?e. Cet exemple suppose qu?une liaison nomm?e `'myBinding'` a ?t? ajout?e au document ? l?aide des m?thodes d?crites plus haut dans cette rubrique.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p125">The following example shows how to use the [Office.select] method to get a [Binding] object promise in a document by specifying its ID in a selector string. It then calls the Binding.[getDataAsync] method to get data from the specified binding. This example assumes that a binding named `'myBinding'` was added to the document using one of the methods described earlier in this topic.</span></span>


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
> <span data-ttu-id="2c0f1-p126">Si la promesse de la m?thode `select` renvoie un objet [Binding], cet objet expose uniquement les quatre m?thodes suivantes de l?objet : [getDataAsync], [setDataAsync], [addHandlerAsync] et [removeHandlerAsync]. Si la promesse ne peut pas renvoyer un objet Binding, le rappel `onError` peut ?tre utilis? pour acc?der ? un objet [asyncResult].error afin d?obtenir plus d?informations. Si vous devez appeler un membre de l?objet Binding autre que les quatre m?thodes expos?es par la promesse d?objet Binding renvoy?e par la m?thode `select`, utilisez plut?t la m?thode [getByIdAsync] en employant la propri?t? [Document.bindings] et la m?thode [Bindings.getByIdAsync] pour r?cup?rer l?objet Binding**.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p126">If the  `select` method promise successfully returns a [Binding] object, that object exposes only the following four methods of the object: [getDataAsync], [setDataAsync], [addHandlerAsync], and [removeHandlerAsync]. If the promise cannot return a  Binding object, the `onError` callback can be used to access an [asyncResult].error object to get more information.If you need to call a member of the Binding object other than the four methods exposed by the Binding object promise returned by the `select` method, instead use the [getByIdAsync] method by using the [Document.bindings] property and Bindings.[getByIdAsync] method to retrieve the Binding** object.</span></span>

## <a name="release-a-binding-by-id"></a><span data-ttu-id="2c0f1-191">Publication d?une liaison par ID</span><span class="sxs-lookup"><span data-stu-id="2c0f1-191">Release a binding by ID</span></span>


<span data-ttu-id="2c0f1-192">L?exemple suivant montre comment utiliser la m?thode [releaseByIdAsync] pour publier une liaison dans un document en sp?cifiant son ID.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-192">The following example shows how use the [releaseByIdAsync] method to release a binding in a document by specifying its ID.</span></span>

```js
Office.context.document.bindings.releaseByIdAsync('myBinding', function (asyncResult) {
    write('Released myBinding!');
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="2c0f1-193">Dans l?exemple, le premier param?tre `id` est l?ID de la liaison ? publier.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-193">In the example, the first `id` parameter is the ID of the binding to release.</span></span>

<span data-ttu-id="2c0f1-p127">La fonction anonyme qui est pass?e dans la fonction comme le deuxi?me param?tre est un rappel qui est ex?cut? lorsque l?op?ration est termin?e. La fonction est appel?e avec un seul param?tre,  [asyncResult], qui contient le statut de l?appel.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p127">The anonymous function that is passed into the function as the second parameter is a callback that is executed when the operation is complete. The function is called with a single parameter,  [asyncResult], which contains the status of the call.</span></span>


## <a name="read-data-from-a-binding"></a><span data-ttu-id="2c0f1-196">Lecture de donn?es ? partir d?une liaison</span><span class="sxs-lookup"><span data-stu-id="2c0f1-196">Read data from a binding</span></span>


<span data-ttu-id="2c0f1-197">L?exemple suivant montre comment utiliser la m?thode [getDataAsync] pour obtenir des donn?es ? partir d?une liaison existante.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-197">The following example shows how to use the [getDataAsync] method to get data from an existing binding.</span></span>


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

 <span data-ttu-id="2c0f1-p128">`myBinding` est une variable qui contient une liaison de texte existante dans le document. Vous pouvez ?galement utiliser [Office.select] pour acc?der ? la liaison avec son identifiant et commencer ? appeler la m?thode [getDataAsync] de la mani?re suivante :</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p128">`myBinding` is a variable that contains an existing text binding in the document. Alternatively, you could use the [Office.select] to access the binding by its ID, and start your call to the [getDataAsync] method, like this:</span></span> 

```js 
Office.select("bindings#myBindingID").getDataAsync
```


<span data-ttu-id="2c0f1-p129">La fonction anonyme qui est pass?e dans la fonction est un rappel qui est ex?cut? lorsque l?op?ration est termin?e. La propri?t? [AsyncResult].value contient les donn?es dans `myBinding`. Le type de valeur d?pend du type de liaison. La liaison dans cet exemple est une liaison de texte. Par cons?quent, la valeur contiendra une cha?ne. Pour obtenir des exemples suppl?mentaires concernant l?utilisation des liaisons de matrice et de tableau, consultez la rubrique sur la m?thode [getDataAsync].</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p129">The anonymous function that is passed into the function is a callback that is executed when the operation is complete. The [AsyncResult].value property contains the data within `myBinding`. The type of the value depends on the binding type. The binding in this example is a text binding. Therefore, the value will contain a string. For additional examples of working with matrix and table bindings, see the [getDataAsync] method topic.</span></span>


## <a name="write-data-to-a-binding"></a><span data-ttu-id="2c0f1-206">?criture de donn?es dans une liaison</span><span class="sxs-lookup"><span data-stu-id="2c0f1-206">Write data to a binding</span></span>

<span data-ttu-id="2c0f1-207">L?exemple suivant montre comment utiliser la m?thode [setDataAsync] pour d?finir des donn?es dans une liaison existante.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-207">The following example shows how to use the [setDataAsync] method to set data in an existing binding.</span></span>

```js
myBinding.setDataAsync('Hello World!', function (asyncResult) { });
```

 <span data-ttu-id="2c0f1-208">`myBinding` est une variable qui contient une liaison de texte existante dans le document.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-208">`myBinding` is a variable that contains an existing text binding in the document.</span></span>

<span data-ttu-id="2c0f1-p130">Dans l?exemple, le premier param?tre est la valeur ? d?finir sur `myBinding`. Comme il s?agit d?une liaison de texte, la valeur est de type `string`. Diff?rents types de liaisons acceptent divers types de donn?es.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p130">In the example, the first parameter is the value to set on  `myBinding`. Because this is a text binding, the value is a `string`. Different binding types accept different types of data.</span></span>

<span data-ttu-id="2c0f1-p131">La fonction anonyme qui est pass?e dans la fonction est un rappel qui est ex?cut? lorsque l?op?ration est termin?e. La fonction est appel?e avec un seul param?tre, `asyncResult`, qui contient l??tat du r?sultat.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p131">The anonymous function that is passed into the function is a callback that is executed when the operation is complete. The function is called with a single parameter,  `asyncResult`, which contains the status of the result.</span></span>

> [!NOTE]
> <span data-ttu-id="2c0f1-214">Depuis la publication d?Excel 2013 SP1 et de la version correspondante d?Excel Online, vous pouvez d?sormais [d?finir la mise en forme lors de l??criture et de la mise ? jour des donn?es dans des tableaux li?s](../excel/excel-add-ins-tables.md).</span><span class="sxs-lookup"><span data-stu-id="2c0f1-214">Starting with the release of the Excel 2013 SP1 and the corresponding build of Excel Online, you can now [set formatting when writing and updating data in bound tables](../excel/excel-add-ins-tables.md).</span></span>


## <a name="detect-changes-to-data-or-the-selection-in-a-binding"></a><span data-ttu-id="2c0f1-215">D?tection des modifications apport?es aux donn?es ou ? la section dans une liaison</span><span class="sxs-lookup"><span data-stu-id="2c0f1-215">Detect changes to data or the selection in a binding</span></span>


<span data-ttu-id="2c0f1-216">L?exemple suivant montre comment lier un gestionnaire d??v?nements ? l??v?nement [DataChanged](https://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedevent) d?une liaison ayant l?ID ? MyBinding ?.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-216">The following example shows how to attach an event handler to the [DataChanged](https://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedevent) event of a binding with an id of "MyBinding".</span></span>


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

<span data-ttu-id="2c0f1-217">est une variable qui contient une liaison de texte existante dans le document.`myBinding`</span><span class="sxs-lookup"><span data-stu-id="2c0f1-217">The `myBinding` is a variable that contains an existing text binding in the document.</span></span>

<span data-ttu-id="2c0f1-p132">Le premier param?tre `eventType` de la m?thode [addHandlerAsync] sp?cifie le nom de l??v?nement auquel s?abonner. [Office.EventType] est une ?num?ration des valeurs de types d??v?nement disponibles. `Office.EventType.BindingDataChanged evaluates to the string `"bindingDataChanged"\`.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p132">The first  `eventType` parameter of the [addHandlerAsync] method specifies the name of the event to subscribe to. [Office.EventType] is an enumeration of available event type values. `Office.EventType.BindingDataChanged evaluates to the string `"bindingDataChanged"\`.</span></span>

<span data-ttu-id="2c0f1-p133">La fonction  `dataChanged` qui est pass?e dans la fonction comme deuxi?me param?tre _handler_ est un gestionnaire d??v?nements qui est ex?cut? lorsque les donn?es dans la liaison sont modifi?es. La fonction est appel?e avec un seul param?tre, _eventArgs_, qui contient une r?f?rence ? la liaison. Cette liaison peut ?tre utilis?e pour r?cup?rer les donn?es mises ? jour.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p133">The  `dataChanged` function that is passed into the function as the second _handler_ parameter is an event handler that is executed when the data in the binding is changed. The function is called with a single parameter, _eventArgs_, which contains a reference to the binding. This binding can be used to retrieve the updated data.</span></span>

<span data-ttu-id="2c0f1-p134">De m?me, vous pouvez d?tecter lorsqu?un utilisateur modifie la s?lection dans une liaison en ajoutant un gestionnaire d??v?nements ? l??v?nement [SelectionChanged] d?une liaison. Pour ce faire, sp?cifiez le param?tre `eventType` de la m?thode [addHandlerAsync] comme `Office.EventType.BindingSelectionChanged` ou `"bindingSelectionChanged"`.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p134">Similarly, you can detect when a user changes selection in a binding by attaching an event handler to the [SelectionChanged] event of a binding. To do that, specify the `eventType` parameter of the [addHandlerAsync] method as `Office.EventType.BindingSelectionChanged` or `"bindingSelectionChanged"`.</span></span>

<span data-ttu-id="2c0f1-p135">Vous pouvez ajouter plusieurs gestionnaires d??v?nements pour un ?v?nement donn? en appelant ? nouveau la m?thode [addHandlerAsync] et en transmettant une fonction de gestionnaire d??v?nements suppl?mentaire pour le param?tre `handler`. Cela fonctionnera correctement tant que le nom de chaque fonction de gestionnaire d??v?nements est unique.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p135">You can add multiple event handlers for a given event by calling the [addHandlerAsync] method again and passing in an additional event handler function for the `handler` parameter. This will work correctly as long as the name of each event handler function is unique.</span></span>


### <a name="remove-an-event-handler"></a><span data-ttu-id="2c0f1-228">Suppression d?un gestionnaire d??v?nements</span><span class="sxs-lookup"><span data-stu-id="2c0f1-228">Remove an event handler</span></span>


<span data-ttu-id="2c0f1-p136">Pour supprimer un gestionnaire d??v?nements pour un ?v?nement, appelez la m?thode [removeHandlerAsync] en transmettant le type d??v?nement en tant que premier param?tre _eventType_, puis le nom de la fonction de gestionnaire d??v?nements ? supprimer comme deuxi?me param?tre _handler_. Par exemple, la fonction suivante supprimera la fonction de gestionnaire d??v?nements `dataChanged` ajout?e dans l?exemple de la section pr?c?dente.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-p136">To remove an event handler for an event, call the [removeHandlerAsync] method passing in the event type as the first _eventType_ parameter, and the name of the event handler function to remove as the second _handler_ parameter. For example, the following function will remove the `dataChanged` event handler function added in the previous section's example.</span></span>


```js
function removeEventHandlerFromBinding() {
    Office.select("bindings#MyBinding").removeHandlerAsync(
        Office.EventType.BindingDataChanged, {handler:dataChanged});
}
```


> [!IMPORTANT]
> <span data-ttu-id="2c0f1-231">Si le param?tre facultatif _handler_ est omis lors de l?appel ? la m?thode [removeHandlerAsync], tous les gestionnaires d??v?nements du param?tre `eventType` sp?cifi? seront supprim?s.</span><span class="sxs-lookup"><span data-stu-id="2c0f1-231">If the optional  _handler_ parameter is omitted when the [removeHandlerAsync] method is called, all event handlers for the specified `eventType` will be removed.</span></span>


## <a name="see-also"></a><span data-ttu-id="2c0f1-232">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="2c0f1-232">See also</span></span>

- [<span data-ttu-id="2c0f1-233">Pr?sentation de l?API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="2c0f1-233">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md) 
- [<span data-ttu-id="2c0f1-234">Programmation asynchrone dans des compl?ments Office</span><span class="sxs-lookup"><span data-stu-id="2c0f1-234">Asynchronous programming in Office Add-ins</span></span>](asynchronous-programming-in-office-add-ins.md)
- [<span data-ttu-id="2c0f1-235">Lecture et ?criture de donn?es dans la s?lection active d?un document ou d?une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="2c0f1-235">Read and write data to the active selection in a document or spreadsheet</span></span>](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
    
[Binding]:               https://dev.office.com/reference/add-ins/shared/binding
[MatrixBinding]:         https://dev.office.com/reference/add-ins/shared/binding.matrixbinding
[TableBinding]:          https://dev.office.com/reference/add-ins/shared/binding.tablebinding
[TextBinding]:           https://dev.office.com/reference/add-ins/shared/binding.textbinding
[getDataAsync]:          https://dev.office.com/reference/add-ins/shared/binding.getdataasync
[setDataAsync]:          https://dev.office.com/reference/add-ins/shared/binding.setdataasync
[SelectionChanged]:      https://dev.office.com/reference/add-ins/shared/binding.bindingselectionchangedevent
[addHandlerAsync]:       https://dev.office.com/reference/add-ins/shared/binding.addhandlerasync
[removeHandlerAsync]:    https://dev.office.com/reference/add-ins/shared/binding.removehandlerasync

[Bindings]:              https://dev.office.com/reference/add-ins/shared/bindings.bindings
[getByIdAsync]:          https://dev.office.com/reference/add-ins/shared/bindings.getbyidasync 
[getAllAsync]:           https://dev.office.com/reference/add-ins/shared/bindings.getallasync
[addFromNamedItemAsync]: https://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync
[addFromSelectionAsync]: https://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync
[addFromPromptAsync]:    https://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync
[releaseByIdAsync]:      https://dev.office.com/reference/add-ins/shared/bindings.releasebyidasync

[AsyncResult]:          https://dev.office.com/reference/add-ins/shared/asyncresult
[Office.BindingType]:   https://dev.office.com/reference/add-ins/shared/bindingtype-enumeration
[Office.select]:        https://dev.office.com/reference/add-ins/shared/office.select 
[Office.EventType]:     https://dev.office.com/reference/add-ins/shared/eventtype-enumeration 
[Document.bindings]:    https://dev.office.com/reference/add-ins/shared/document.bindings


[TableBinding.rowCount]: https://dev.office.com/reference/add-ins/shared/binding.tablebinding.rowcount
[BindingSelectionChangedEventArgs]: https://dev.office.com/reference/add-ins/shared/binding.bindingselectionchangedeventargs
