<span data-ttu-id="a73fa-101">Dans cette étape du didacticiel, vous allez ajouter un autre bouton au ruban qui, lorsque l’utilisateur clique dessus, exécute une fonction qui vous allez définir et qui active/désactive la protection de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="a73fa-101">In this step of the tutorial, you'll add another button to the ribbon that, when chosen, executes a function that you'll define to toggle worksheet protection on and off.</span></span>

> [!NOTE]
> <span data-ttu-id="a73fa-102">Cette page décrit une étape individuelle du didacticiel sur le complément Excel.</span><span class="sxs-lookup"><span data-stu-id="a73fa-102">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="a73fa-103">Si vous êtes arrivé à cette page via les résultats du moteur de recherche ou d’un autre lien direct, accédez à la page d’introduction du [didacticiel sur le complément Excel](../tutorials/excel-tutorial.yml) pour démarrer le didacticiel à partir du début.</span><span class="sxs-lookup"><span data-stu-id="a73fa-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="configure-the-manifest-to-add-a-second-ribbon-button"></a><span data-ttu-id="a73fa-104">Configuration du manifeste pour ajouter un deuxième bouton de ruban</span><span class="sxs-lookup"><span data-stu-id="a73fa-104">Configure the manifest to add a second ribbon button</span></span>

1. <span data-ttu-id="a73fa-105">Ouvrez le fichier manifeste **my-office-add-in-manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="a73fa-105">Open the manifest file **my-office-add-in-manifest.xml**.</span></span>
2. <span data-ttu-id="a73fa-106">Recherchez l’élément `<Control>`.</span><span class="sxs-lookup"><span data-stu-id="a73fa-106">Find the `<Control>` element.</span></span> <span data-ttu-id="a73fa-107">Cet élément définit le bouton **Afficher le volet des pages** sur le ruban **Accueil** que vous utilisez pour lancer le complément.</span><span class="sxs-lookup"><span data-stu-id="a73fa-107">This element defines the **Show Taskpane** button on the **Home** ribbon you have been using to launch the add-in.</span></span> <span data-ttu-id="a73fa-108">Nous allons ajouter un deuxième bouton au même groupe sur le ruban **Accueil**.</span><span class="sxs-lookup"><span data-stu-id="a73fa-108">We're going to add a second button to the same group on the **Home** ribbon.</span></span> <span data-ttu-id="a73fa-109">Entre la balise Control de fin (`</Control>`) et la balise Group de fin (`</Group>`), ajoutez le balisage suivant.</span><span class="sxs-lookup"><span data-stu-id="a73fa-109">In between the end Control tag (`</Control>`) and the end Group tag (`</Group>`), add the following markup.</span></span>

    ```xml
    <Control xsi:type="Button" id="<!--TODO1: Unique (in manifest) name for button -->">
        <Label resid="<!--TODO2: Button label -->" />
        <Supertip>            
            <Title resid="<!-- TODO3: Button tool tip title -->" />
            <Description resid="<!-- TODO4: Button tool tip description -->" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Contoso.tpicon_16x16" />
            <bt:Image size="32" resid="Contoso.tpicon_32x32" />
            <bt:Image size="80" resid="Contoso.tpicon_80x80" />
        </Icon>
        <Action xsi:type="<!-- TODO5: Specify the type of action-->">
            <!-- TODO6: Identify the function.-->
        </Action>
    </Control>
    ```

3. <span data-ttu-id="a73fa-110">Remplacez `TODO1` par une chaîne qui attribue un ID unique au bouton au sein de ce fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="a73fa-110">Replace `TODO1` with a string that gives the button an ID that is unique within this manifest file.</span></span> <span data-ttu-id="a73fa-111">Il n’existe qu’un seul autre bouton dans le fichier manifeste, donc ce n’est pas difficile.</span><span class="sxs-lookup"><span data-stu-id="a73fa-111">There's only one other button in the manifest, so this isn't difficult.</span></span> <span data-ttu-id="a73fa-112">Étant donné que notre bouton va activer ou désactiver la protection de la feuille de calcul, utilisez « ToggleProtection ».</span><span class="sxs-lookup"><span data-stu-id="a73fa-112">Since our button is going to toggle protection of the worksheet on and off, use "ToggleProtection".</span></span> <span data-ttu-id="a73fa-113">Lorsque vous avez terminé, la balise Control de début doit ressembler à ce qui suit :</span><span class="sxs-lookup"><span data-stu-id="a73fa-113">When you are done, the entire start Control tag should look like the following:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
    ```

4. <span data-ttu-id="a73fa-114">Les trois éléments `TODO` suivants définissent les éléments « resid », c’est-à-dire les ID de ressource.</span><span class="sxs-lookup"><span data-stu-id="a73fa-114">The next three `TODO`s set "resid"s, which is short for resource ID.</span></span> <span data-ttu-id="a73fa-115">Une ressource est une chaîne. Vous allez créer ces trois chaînes lors d’une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="a73fa-115">A resource is a string, and you'll create these three strings in a later step.</span></span> <span data-ttu-id="a73fa-116">Pour l’instant, vous devez attribuer des ID aux ressources.</span><span class="sxs-lookup"><span data-stu-id="a73fa-116">For now, you need to give IDs to the resources.</span></span> <span data-ttu-id="a73fa-117">L’étiquette du bouton doit indiquer « Toggle Protection », mais l’*ID* de cette chaîne doit être « ProtectionButtonLabel », donc l’élément `Label` terminé doit ressembler au code suivant :</span><span class="sxs-lookup"><span data-stu-id="a73fa-117">The button label should read "Toggle Protection", but the *ID* of this string should be "ProtectionButtonLabel", so the completed `Label` element should look like the following code:</span></span>

    ```xml
    <Label resid="ProtectionButtonLabel" />
    ```

5. <span data-ttu-id="a73fa-118">L’élément `SuperTip` définit l’info-bulle du bouton.</span><span class="sxs-lookup"><span data-stu-id="a73fa-118">The `SuperTip` element defines the tool tip for the button.</span></span> <span data-ttu-id="a73fa-119">Le titre de l’info-bulle doit être identique à l’étiquette du bouton, nous utilisons donc le même ID de ressource : « ProtectionButtonLabel ».</span><span class="sxs-lookup"><span data-stu-id="a73fa-119">The tool tip title should be the same as the button label, so we use the very same resource ID: "ProtectionButtonLabel".</span></span> <span data-ttu-id="a73fa-120">La description de l’info-bulle sera « Cliquez pour activer/désactiver la protection de la feuille de calcul ».</span><span class="sxs-lookup"><span data-stu-id="a73fa-120">The tool tip description will be "Click to turn protection of the worksheet on and off".</span></span> <span data-ttu-id="a73fa-121">Néanmoins, l’élément `ID` doit être « ProtectionButtonToolTip ».</span><span class="sxs-lookup"><span data-stu-id="a73fa-121">But the `ID` should be "ProtectionButtonToolTip".</span></span> <span data-ttu-id="a73fa-122">Ainsi, lorsque vous avez terminé, l’ensemble du balisage `SuperTip` doit ressembler au code suivant :</span><span class="sxs-lookup"><span data-stu-id="a73fa-122">So, when you are done, the whole `SuperTip` markup should look like the following code:</span></span> 

    ```xml
    <Supertip>            
        <Title resid="ProtectionButtonLabel" />
        <Description resid="ProtectionButtonToolTip" />
    </Supertip>
    ```

   > [!NOTE] 
   > <span data-ttu-id="a73fa-123">Dans un complément de production, vous n’utiliseriez pas la même icône pour deux boutons différents, mais pour simplifier ce didacticiel, nous allons le faire.</span><span class="sxs-lookup"><span data-stu-id="a73fa-123">In a production add-in, you would not want to use the same icon for two different buttons; but to simplify this tutorial, we'll do that.</span></span> <span data-ttu-id="a73fa-124">Par conséquent, le balisage `Icon` de notre nouvel élément `Control` est simplement une copie de l’élément `Icon` provenant de l’élément `Control` existant.</span><span class="sxs-lookup"><span data-stu-id="a73fa-124">So the `Icon` markup in our new `Control` is just a copy of the `Icon` element from the existing `Control`.</span></span> 

6. <span data-ttu-id="a73fa-125">Le type de l’élément `Action` se trouvant à l’intérieur de l’élément `Control` d’origine qui était déjà présent dans le fichier manifeste est défini sur `ShowTaskpane`, mais notre nouveau bouton ne va pas ouvrir un volet Office, il va exécuter une fonction personnalisée que vous allez créer à une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="a73fa-125">The `Action` element inside the original `Control` element that was already present in the manifest, has its type set to `ShowTaskpane`, but our new button isn't going to open a task pane; it's going to run a custom function that you create in a later step.</span></span> <span data-ttu-id="a73fa-126">Il faut donc remplacer `TODO5` par `ExecuteFunction`, c’est-à-dire le type d’action pour les boutons qui déclenchent des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="a73fa-126">So replace `TODO5` with `ExecuteFunction` which is the action type for buttons that trigger custom functions.</span></span> <span data-ttu-id="a73fa-127">La balise `Action` de début doit ressembler au code suivant :</span><span class="sxs-lookup"><span data-stu-id="a73fa-127">The start `Action` tag should look like the following code:</span></span>
 
    ```xml
    <Action xsi:type="ExecuteFunction">
    ```

7. <span data-ttu-id="a73fa-128">L’élément `Action` d’origine possède des éléments enfants qui spécifient un ID de volet Office ainsi qu’une URL de la page qui doit être ouverte dans le volet Office.</span><span class="sxs-lookup"><span data-stu-id="a73fa-128">The original `Action` element has child elements that specify a task pane ID and a URL of the page that should be opened in the task pane.</span></span> <span data-ttu-id="a73fa-129">Toutefois, un élément `Action` de type `ExecuteFunction` comporte un élément enfant unique qui nomme la fonction que le contrôle exécute.</span><span class="sxs-lookup"><span data-stu-id="a73fa-129">But an `Action` element of the `ExecuteFunction` type has a single child element that names the function that the control executes.</span></span> <span data-ttu-id="a73fa-130">Vous créerez cette fonction à une étape ultérieure, et la nommerez `toggleProtection`.</span><span class="sxs-lookup"><span data-stu-id="a73fa-130">You'll create that function in a later step, and it will be called `toggleProtection`.</span></span> <span data-ttu-id="a73fa-131">Par conséquent, remplacez `TODO6` par le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="a73fa-131">So, replace `TODO6` with the following markup:</span></span>
 
    ```xml
    <FunctionName>toggleProtection</FunctionName>
    ```

    <span data-ttu-id="a73fa-132">Le balisage `Control` complet doit à présent ressembler à ce qui suit :</span><span class="sxs-lookup"><span data-stu-id="a73fa-132">The entire `Control` markup should now look like the following:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
        <Label resid="ProtectionButtonLabel" />
        <Supertip>            
            <Title resid="ProtectionButtonLabel" />
            <Description resid="ProtectionButtonToolTip" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Contoso.tpicon_16x16" />
            <bt:Image size="32" resid="Contoso.tpicon_32x32" />
            <bt:Image size="80" resid="Contoso.tpicon_80x80" />
        </Icon>
        <Action xsi:type="ExecuteFunction">
           <FunctionName>toggleProtection</FunctionName>
        </Action>
    </Control>
    ```

8. <span data-ttu-id="a73fa-133">Faites défiler vers le bas jusqu’à la section `Resources` du manifeste.</span><span class="sxs-lookup"><span data-stu-id="a73fa-133">Scroll down to the `Resources` section of the manifest.</span></span>

9. <span data-ttu-id="a73fa-134">Ajoutez le balisage suivant en tant qu’enfant de l’élément `bt:ShortStrings`.</span><span class="sxs-lookup"><span data-stu-id="a73fa-134">Add the following markup as a child of the `bt:ShortStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonLabel" DefaultValue="Toggle Worksheet Protection" />
    ```

10. <span data-ttu-id="a73fa-135">Ajoutez le balisage suivant en tant qu’enfant de l’élément `bt:LongStrings`.</span><span class="sxs-lookup"><span data-stu-id="a73fa-135">Add the following markup as a child of the `bt:LongStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonToolTip" DefaultValue="Click to protect or unprotect the current worksheet." />
    ```

11. <span data-ttu-id="a73fa-136">Enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="a73fa-136">Be sure to save the file.</span></span>

## <a name="create-the-function-that-protects-the-sheet"></a><span data-ttu-id="a73fa-137">Création de la fonction qui protège la feuille</span><span class="sxs-lookup"><span data-stu-id="a73fa-137">Create the function that protects the sheet</span></span>

1. <span data-ttu-id="a73fa-138">Ouvrez le fichier \function-file\function-file.js.</span><span class="sxs-lookup"><span data-stu-id="a73fa-138">Open the file \function-file\function-file.js.</span></span>

2. <span data-ttu-id="a73fa-139">Le fichier possède déjà une expression de fonction appelée immédiatement (IIFE).</span><span class="sxs-lookup"><span data-stu-id="a73fa-139">The file already has an Immediately Invoked Function Expression (IFFE).</span></span> <span data-ttu-id="a73fa-140">Aucune logique d’initialisation personnalisée n’est nécessaire, donc laissez la fonction qui a été attribuée à `Office.initialize` avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="a73fa-140">No custom initialization logic is needed, so leave the function that is assigned to `Office.initialize` with an empty body.</span></span> <span data-ttu-id="a73fa-141">(Mais ne la supprimez pas.</span><span class="sxs-lookup"><span data-stu-id="a73fa-141">(But do not delete it.</span></span> <span data-ttu-id="a73fa-142">La propriété `Office.initialize` ne peut pas être null ou non définie.) *En dehors de l’IIFE*, ajoutez le code suivant.</span><span class="sxs-lookup"><span data-stu-id="a73fa-142">The `Office.initialize` property cannot be null or undefined.) *Outside of the IIFE*, add the following code.</span></span> <span data-ttu-id="a73fa-143">Notez que nous spécifions un paramètre `args` pour la méthode et que la toute dernière ligne de la méthode appelle `args.completed`.</span><span class="sxs-lookup"><span data-stu-id="a73fa-143">Note that we specify an `args` parameter to the method and the very last line of the method calls `args.completed`.</span></span> <span data-ttu-id="a73fa-144">Il s’agit d’une condition requise pour toutes les commandes de type **ExecuteFunction**.</span><span class="sxs-lookup"><span data-stu-id="a73fa-144">This is a requirement for all add-in commands of type **ExecuteFunction**.</span></span> <span data-ttu-id="a73fa-145">Elle signale à l’application hôte Office que la fonction est terminée et que l’interface utilisateur est à nouveau réactive.</span><span class="sxs-lookup"><span data-stu-id="a73fa-145">It signals the Office host application that the function has finished and the UI can become responsive again.</span></span>

    ```javascript
    function toggleProtection(args) {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to reverse the protection status of the current worksheet.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
        args.completed();
    }
    ```

3. <span data-ttu-id="a73fa-146">Remplacez `TODO1` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="a73fa-146">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="a73fa-147">Ce code utilise la propriété de protection de l’objet de feuille de calcul dans un modèle de bouton bascule standard.</span><span class="sxs-lookup"><span data-stu-id="a73fa-147">This code uses the worksheet object's protection property in a standard toggle pattern.</span></span> <span data-ttu-id="a73fa-148">L’élément `TODO2` sera expliqué dans la section suivante.</span><span class="sxs-lookup"><span data-stu-id="a73fa-148">The `TODO2` will be explained in the next section.</span></span>

    ```javascript
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // TODO2: Queue command to load the sheet's "protection.protected" property from
    //        the document and re-synchronize the document and task pane.

     if (sheet.protection.protected) {
        sheet.protection.unprotect();
    } else {
        sheet.protection.protect();
    }
    ``` 

## <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a><span data-ttu-id="a73fa-149">Ajoutez du code pour récupérer des propriétés de document dans les objets de script du volet Office</span><span class="sxs-lookup"><span data-stu-id="a73fa-149">Add code to fetch document properties into the task pane's script objects</span></span>

<span data-ttu-id="a73fa-150">Dans toutes les fonctions précédentes de cette série de didacticiels, vous avez mis en file d’attente des commandes pour écrire (*write*) dans le document Office.</span><span class="sxs-lookup"><span data-stu-id="a73fa-150">In all the earlier functions in this series of tutorials, you queued commands to *write* to the Office document.</span></span> <span data-ttu-id="a73fa-151">Chaque fonction se terminait par un appel de la méthode `context.sync()` qui envoie les commandes en file d’attente au document pour qu’elles soient exécutées.</span><span class="sxs-lookup"><span data-stu-id="a73fa-151">Each function ended with a call to the `context.sync()` method which sends the queued commands to the document to be executed.</span></span> <span data-ttu-id="a73fa-152">Cependant, le code que vous avez ajouté dans la dernière étape appelle la propriété `sheet.protection.protected` et c’est une différence significative par rapport aux fonctions antérieures que vous avez écrites, car l’objet `sheet` est uniquement un objet de proxy qui existe dans le script de votre volet Office.</span><span class="sxs-lookup"><span data-stu-id="a73fa-152">But the code you added in the last step calls the `sheet.protection.protected` property, and this is a significant difference from the earlier functions you wrote, because the `sheet` object is only a proxy object that exists in your task pane's script.</span></span> <span data-ttu-id="a73fa-153">Il ne connaît pas l’état de protection réel du document, donc sa propriété `protection.protected` ne peut pas contenir une valeur réelle.</span><span class="sxs-lookup"><span data-stu-id="a73fa-153">It doesn't know what the actual protection state of the document is, so its `protection.protected` property can't have a real value.</span></span> <span data-ttu-id="a73fa-154">Tout d’abord, il faut récupérer l’état de protection dans le document et l’utiliser pour définir la valeur de `sheet.protection.protected`.</span><span class="sxs-lookup"><span data-stu-id="a73fa-154">It is necessary to first fetch the protection status from the document and use it set the value of `sheet.protection.protected`.</span></span> <span data-ttu-id="a73fa-155">Seulement ensuite, la propriété `sheet.protection.protected` peut être appelée sans générer d’exception.</span><span class="sxs-lookup"><span data-stu-id="a73fa-155">Only then can `sheet.protection.protected` be called without causing an exception to be thrown.</span></span> <span data-ttu-id="a73fa-156">Ce processus de récupération comporte trois étapes :</span><span class="sxs-lookup"><span data-stu-id="a73fa-156">This fetching process has three steps:</span></span>

   1. <span data-ttu-id="a73fa-157">Mettez en file d’attente une commande de chargement (c’est-à-dire, fetch) des propriétés que votre code doit lire.</span><span class="sxs-lookup"><span data-stu-id="a73fa-157">Queue a command to load (that is; fetch) the properties that your code needs to read.</span></span>
   2. <span data-ttu-id="a73fa-158">Appelez la méthode `sync` de l’objet de contexte pour envoyer la commande mise en file d’attente vers le document pour exécution, et renvoyez les informations demandées.</span><span class="sxs-lookup"><span data-stu-id="a73fa-158">Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.</span></span>
   3. <span data-ttu-id="a73fa-159">Étant donné que la méthode `sync` est asynchrone, assurez-vous qu’elle est terminée avant que votre code appelle les propriétés qui ont été récupérées.</span><span class="sxs-lookup"><span data-stu-id="a73fa-159">Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.</span></span>

<span data-ttu-id="a73fa-160">Ces étapes doivent être effectuées à chaque fois que votre code doit lire (*read*) des informations provenant du document Office.</span><span class="sxs-lookup"><span data-stu-id="a73fa-160">These steps must be completed whenever your code needs to *read* information from the Office document.</span></span>

1. <span data-ttu-id="a73fa-p112">Dans la fonction `toggleProtection`, remplacez `TODO2` par le code suivant. Remarque :</span><span class="sxs-lookup"><span data-stu-id="a73fa-p112">In the `toggleProtection` function, replace `TODO2` with the following code. Note:</span></span>
   - <span data-ttu-id="a73fa-163">Chaque objet Excel possède une méthode `load`.</span><span class="sxs-lookup"><span data-stu-id="a73fa-163">Every Excel object has a `load` method.</span></span> <span data-ttu-id="a73fa-164">Vous spécifiez les propriétés de l’objet que vous voulez lire dans le paramètre en tant que chaîne de noms séparés par des virgules.</span><span class="sxs-lookup"><span data-stu-id="a73fa-164">You specify the properties of the object that you want to read in the parameter as a string of comma-delimited names.</span></span> <span data-ttu-id="a73fa-165">Dans ce cas, la propriété que vous devez lire est une sous-propriété de la propriété `protection`.</span><span class="sxs-lookup"><span data-stu-id="a73fa-165">In this case, the property you need to read is a subproperty of the `protection` property.</span></span> <span data-ttu-id="a73fa-166">Pour référence la sous-propriété, procédez presque exactement de la même façon que vous le feriez à n’importe quel autre emplacement de votre code, sauf que vous devez utiliser une barre oblique (« / ») au lieu d’un point « . ».</span><span class="sxs-lookup"><span data-stu-id="a73fa-166">You reference the subproperty almost exactly as you would anywhere else in your code, with the exception that you use a forward slash ('/') character instead of a "." character.</span></span>
   - <span data-ttu-id="a73fa-167">Pour être sûr que la logique de bouton bascule, qui lit `sheet.protection.protected`, ne s’exécute pas tant que la synchronisation (`sync`) n’est pas terminée et que l’élément `sheet.protection.protected` n’a pas été affecté à la valeur correcte récupérée à partir du document, elle sera déplacée (à l’étape suivante) dans une fonction `then` qui ne s’exécutera pas tant que la synchronisation (`sync`) ne sera pas terminée.</span><span class="sxs-lookup"><span data-stu-id="a73fa-167">To ensure that the toggle logic, which reads `sheet.protection.protected`, does not run until after the `sync` is complete and the `sheet.protection.protected` has been assigned the correct value that is fetched from the document, it will be moved (in the next step) into a `then` function that won't run until the `sync` has completed.</span></span> 

    ```javascript
    sheet.load('protection/protected');
    return context.sync()
        .then(
            function() {
                // TODO3: Move the queued toggle logic here.
            }
        )
        // TODO4: Move the final call of `context.sync` here and ensure that it
        //        does not run until the toggle logic has been queued.
    ``` 

2. <span data-ttu-id="a73fa-168">Il n’est pas possible que deux instructions `return` se trouvent dans le même chemin de code, donc supprimez la dernière ligne `return context.sync();` à la fin de la fonction `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="a73fa-168">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Excel.run`.</span></span> <span data-ttu-id="a73fa-169">Vous ajouterez un nouvel élément final `context.sync` dans une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="a73fa-169">You will add a new final `context.sync`, in a later step.</span></span>
3. <span data-ttu-id="a73fa-170">Coupez la structurer `if ... else` dans la fonction `toggleProtection` et collez-la à la place de `TODO3`.</span><span class="sxs-lookup"><span data-stu-id="a73fa-170">Cut the `if ... else` structure in the `toggleProtection` function and paste it in place of `TODO3`.</span></span>
4. <span data-ttu-id="a73fa-p115">Remplacez `TODO4` par le code suivant. Veuillez noter les informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="a73fa-p115">Replace `TODO4` with the following code. Note:</span></span>
   - <span data-ttu-id="a73fa-173">Le fait de transmettre la méthode `sync` à une fonction `then` permet de s’assurer qu’elle n’est pas exécutée tant que `sheet.protection.unprotect()` ou `sheet.protection.protect()` n’a pas été mis en file d’attente.</span><span class="sxs-lookup"><span data-stu-id="a73fa-173">Passing the `sync` method to a `then` function ensures that it does not run until either `sheet.protection.unprotect()` or `sheet.protection.protect()` has been queued.</span></span>
   - <span data-ttu-id="a73fa-174">La méthode `then` appelle n’importe quelle fonction qui lui est transmise, et vous ne souhaitez pas appeler `sync` deux fois, donc omettez les parenthèses « () » à la fin de `context.sync`.</span><span class="sxs-lookup"><span data-stu-id="a73fa-174">The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so leave off the "()" from the end of `context.sync`.</span></span>

    ```javascript
    .then(context.sync);
    ```

   <span data-ttu-id="a73fa-175">Lorsque vous avez terminé, la fonction entière doit ressembler à ce qui suit :</span><span class="sxs-lookup"><span data-stu-id="a73fa-175">When you are done, the entire function should look like the following:</span></span>

    ```javascript
    function toggleProtection(args) {
        Excel.run(function (context) {            
          const sheet = context.workbook.worksheets.getActiveWorksheet();          
          sheet.load('protection/protected');

          return context.sync()
              .then(
                  function() {
                    if (sheet.protection.protected) {
                        sheet.protection.unprotect();
                    } else {
                        sheet.protection.protect();
                    }
                  }
              )
              .then(context.sync);
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
        args.completed();
    }
    ```


## <a name="configure-the-script-loading-html-file"></a><span data-ttu-id="a73fa-176">Configuration du fichier HTML de chargement de script</span><span class="sxs-lookup"><span data-stu-id="a73fa-176">Configure the script-loading HTML file</span></span>

<span data-ttu-id="a73fa-177">Ouvrez le fichier /function-file/function-file.html.</span><span class="sxs-lookup"><span data-stu-id="a73fa-177">Open the /function-file/function-file.html file.</span></span> <span data-ttu-id="a73fa-178">Il s’agit d’un fichier HTML sans interface utilisateur qui est appelé lorsque l’utilisateur appuie sur le bouton **Toggle Worksheet Protection**.</span><span class="sxs-lookup"><span data-stu-id="a73fa-178">This is a UI-less HTML file that is called when the user presses the **Toggle Worksheet Protection** button.</span></span> <span data-ttu-id="a73fa-179">Son objectif consiste à charger la méthode JavaScript qui doit s’exécuter lorsque l’utilisateur appuie sur le bouton.</span><span class="sxs-lookup"><span data-stu-id="a73fa-179">Its purpose is to load the JavaScript method that should run when the button is pushed.</span></span> <span data-ttu-id="a73fa-180">Vous n’allez pas modifier ce fichier.</span><span class="sxs-lookup"><span data-stu-id="a73fa-180">You are not going to change this file.</span></span> <span data-ttu-id="a73fa-181">Remarquez simplement que la deuxième balise `<script>` charge le fichier functionfile.js.</span><span class="sxs-lookup"><span data-stu-id="a73fa-181">Simply note that the second `<script>` tag loads the functionfile.js.</span></span>

   > [!NOTE]
   > <span data-ttu-id="a73fa-182">Le fichier function-file.html et le fichier function-file.js qu’il charge s’exécutent dans un processus Internet Explorer entièrement distinct dans le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="a73fa-182">The function-file.html file and the function-file.js file that it loads run in an entirely separate IE process from the add-in's task pane.</span></span> <span data-ttu-id="a73fa-183">Si le fichier function-file.js était transpilé dans le même fichier bundle.js en tant que fichier app.js, le complément devrait charger deux copies du fichier bundle.js, ce qui irait à l’encontre l’objectif de groupement.</span><span class="sxs-lookup"><span data-stu-id="a73fa-183">If the function-file.js was transpiled into the same bundle.js file as the app.js file, then the add-in would have to load two copies of the bundle.js file, which defeats the purpose of bundling.</span></span> <span data-ttu-id="a73fa-184">En outre, le fichier function-file.js ne contient pas de code JavaScript car Internet Explorer ne prend pas en charge ce type de code.</span><span class="sxs-lookup"><span data-stu-id="a73fa-184">In addition, the function-file.js file does not contain any JavaScript that is unsupported by IE.</span></span> <span data-ttu-id="a73fa-185">C’est pour ces deux raisons que ce complément ne transpile pas le fichier function-file.js du tout.</span><span class="sxs-lookup"><span data-stu-id="a73fa-185">For these two reasons, this add-in does not transpile the function-file.js at all.</span></span> 

## <a name="test-the-add-in"></a><span data-ttu-id="a73fa-186">Test du complément</span><span class="sxs-lookup"><span data-stu-id="a73fa-186">Test the add-in</span></span>

1. <span data-ttu-id="a73fa-187">Fermez toutes les applications Office, y compris Excel.</span><span class="sxs-lookup"><span data-stu-id="a73fa-187">Close all Office applications, including Excel.</span></span> 
2. <span data-ttu-id="a73fa-188">Supprimez le cache Office en supprimant le contenu du dossier de cache.</span><span class="sxs-lookup"><span data-stu-id="a73fa-188">Delete the Office cache by deleting the contents of the cache folder.</span></span> <span data-ttu-id="a73fa-189">Cette opération est nécessaire pour effacer complètement de l’hôte l’ancienne version du complément.</span><span class="sxs-lookup"><span data-stu-id="a73fa-189">This is necessary to completely clear the old version of the add-in from the host.</span></span> 
    - <span data-ttu-id="a73fa-190">Pour Windows : `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="a73fa-190">For Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>
    - <span data-ttu-id="a73fa-191">Pour Mac : `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="a73fa-191">For Mac: `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span>
3. <span data-ttu-id="a73fa-192">Si, pour une quelconque raison, votre serveur n’est pas en cours d’exécution, accédez au dossier **Start** du projet et exécutez la commande `npm start` dans une fenêtre Git Bash ou une invite système Node.JS.</span><span class="sxs-lookup"><span data-stu-id="a73fa-192">If for any reason, your server is not running, then in a Git Bash window, or Node.JS-enabled system prompt, navigate to the **Start** folder of the project and run the command `npm start`.</span></span> <span data-ttu-id="a73fa-193">Vous n’avez pas besoin de recréer le projet, car le seul fichier JavaScript que vous avez modifié ne fait pas partie du fichier bundle.js créé.</span><span class="sxs-lookup"><span data-stu-id="a73fa-193">You do not need to rebuild the project because the only JavaScript file you changed is not part of the built bundle.js.</span></span>
4. <span data-ttu-id="a73fa-194">À l’aide de la nouvelle version du fichier manifeste modifié, répétez le processus de chargement de version test en utilisant l’une des méthodes suivantes.</span><span class="sxs-lookup"><span data-stu-id="a73fa-194">Using the new version of the changed manifest file, repeat the sideloading process by using one of the following methods.</span></span> <span data-ttu-id="a73fa-195">*Vous devez remplacer la copie précédente du fichier manifeste.*</span><span class="sxs-lookup"><span data-stu-id="a73fa-195">*You should overwrite the previous copy of the manifest file.*</span></span>
    - <span data-ttu-id="a73fa-196">Windows : [Chargement de version test des compléments Office](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="a73fa-196">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="a73fa-197">Excel Online : [Chargement de versions test des compléments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="a73fa-197">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="a73fa-198">iPad et Mac : [Chargement de version test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="a73fa-198">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>
7. <span data-ttu-id="a73fa-199">Ouvrez une feuille de calcul dans Excel.</span><span class="sxs-lookup"><span data-stu-id="a73fa-199">Open any worksheet in Excel.</span></span>
8. <span data-ttu-id="a73fa-p121">Sur le ruban **Accueil**, sélectionnez **Toggle Worksheet Protection** (Activer/Désactiver la protection de la feuille de calcul). Notez que la plupart des contrôles figurant sur le ruban sont désactivés (et visuellement grisés) comme illustré dans la capture d’écran ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="a73fa-p121">On the **Home** ribbon, choose **Toggle Worksheet Protection**. Note that most of the controls on the ribbon are disabled (and visually grayed-out) as seen in screenshot below.</span></span> 
9. <span data-ttu-id="a73fa-202">Sélectionnez une cellule comme vous le feriez si vous vouliez modifier son contenu.</span><span class="sxs-lookup"><span data-stu-id="a73fa-202">Choose a cell as you would if you wanted to change its content.</span></span> <span data-ttu-id="a73fa-203">Vous rencontrez une erreur indiquant que la feuille de calcul est protégée.</span><span class="sxs-lookup"><span data-stu-id="a73fa-203">You get an error telling you that the worksheet is protected.</span></span>
10. <span data-ttu-id="a73fa-204">Sélectionnez **Toggle Worksheet Protection** à nouveau pour réactiver les contrôles. Vous pouvez alors modifier une nouvelle fois les valeurs de cellule.</span><span class="sxs-lookup"><span data-stu-id="a73fa-204">Choose **Toggle Worksheet Protection** again, and the controls are reenabled, and you can change cell values again.</span></span>

    ![Didacticiel Excel - Ruban avec protection activée](../images/excel-tutorial-ribbon-with-protection-on.png)
