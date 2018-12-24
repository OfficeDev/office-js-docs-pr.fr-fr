# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="474ea-101">Didacticiel : créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="474ea-101">Tutorial: Create custom functions in Excel</span></span>

## <a name="introduction"></a><span data-ttu-id="474ea-102">Présentation</span><span class="sxs-lookup"><span data-stu-id="474ea-102">Introduction</span></span>

<span data-ttu-id="474ea-103">Les fonctions personnalisées vous permettent d’ajouter de nouvelles fonctions dans Excel en définissant ces fonctions dans JavaScript dans le cadre d’un complément.</span><span class="sxs-lookup"><span data-stu-id="474ea-103">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="474ea-104">Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe quelle fonction native d’Excel, telle que `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="474ea-104">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="474ea-105">Vous pouvez créer des fonctions personnalisées qui effectuent des tâches simples telles que des calculs personnalisés ou des tâches plus complexes telles que la diffusion en continu des données en temps réel à partir du web dans une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="474ea-105">You can create custom functions that perform simple tasks such as custom calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="474ea-106">Dans ce didacticiel, vous allez :</span><span class="sxs-lookup"><span data-stu-id="474ea-106">In this tutorial, you will:</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="474ea-107">Créer un projet de fonctions personnalisées à l’aide du générateur Yo Office</span><span class="sxs-lookup"><span data-stu-id="474ea-107">Create a custom functions project by using the Yo Office generator</span></span>
> * <span data-ttu-id="474ea-108">Utiliser une fonction personnalisée prédéfinie pour effectuer un calcul simple</span><span class="sxs-lookup"><span data-stu-id="474ea-108">Use a prebuilt custom function to perform a simple calculation</span></span>
> * <span data-ttu-id="474ea-109">Créer une fonction personnalisée qui demande les données à partir du web</span><span class="sxs-lookup"><span data-stu-id="474ea-109">Create a custom function that requests data from the web</span></span>
> * <span data-ttu-id="474ea-110">Créer une fonction personnalisée qui diffuse les données en temps réel à partir du web</span><span class="sxs-lookup"><span data-stu-id="474ea-110">Create a custom function that streams real-time data from the web</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a><span data-ttu-id="474ea-111">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="474ea-111">Prerequisites</span></span>

* <span data-ttu-id="474ea-112">[Node.js](https://nodejs.org/en/) (version 8.0.0 ou ultérieure)</span><span class="sxs-lookup"><span data-stu-id="474ea-112">[Node.js](https://nodejs.org/en/) (version 8.0.0 or later)</span></span>

* <span data-ttu-id="474ea-113">[Git Bash](https://git-scm.com/downloads) (ou un autre client Git)</span><span class="sxs-lookup"><span data-stu-id="474ea-113">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

* <span data-ttu-id="474ea-114">La dernière version de[Yeoman](https://yeoman.io/) et de [Yeoman Générateur de compléments Office](https://www.npmjs.com/package/generator-office). Pour installer ces outils globalement, exécutez la commande suivante à partir de l’invite de commande :</span><span class="sxs-lookup"><span data-stu-id="474ea-114">The latest version of [Yeoman](https://yeoman.io/) and the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command from the command prompt:</span></span>

    ```
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > <span data-ttu-id="474ea-115">Même si vous avez précédemment installé la Yeoman générateur, nous vous recommandons une mise à jour de votre package à partir de la dernière version de npm.</span><span class="sxs-lookup"><span data-stu-id="474ea-115">Even if you have previously installed the Yeoman generator, we recommend updating your package to the latest version from npm.</span></span>

* <span data-ttu-id="474ea-116">Excel pour Windows (version 64 bits 1810 ou ultérieure) ou Excel Online</span><span class="sxs-lookup"><span data-stu-id="474ea-116">Excel for Windows (version 1810 or later) or Excel Online</span></span>

* <span data-ttu-id="474ea-117">Rejoignez le[programme Office Insider](https://products.office.com/office-insider)(\*\* niveau\*\*Insider, anciennement appelé « Insider Fast »)</span><span class="sxs-lookup"><span data-stu-id="474ea-117">Join the [Office Insider program](https://products.office.com/office-insider) (**Insider** level -- formerly called "Insider Fast")</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="474ea-118">Créer un projet de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="474ea-118">Create a custom functions project</span></span>

 <span data-ttu-id="474ea-119">Pour commencer, vous utiliserez le Yeoman Générateur pour créer le projet de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="474ea-119">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="474ea-120">Cette option définit votre projet, avec la structure de dossiers correct, les fichiers source et les dépendances pour commencer le codage de vos fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="474ea-120">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. <span data-ttu-id="474ea-121">Exécutez la commande suivante, puis répondez aux invitations comme suit.</span><span class="sxs-lookup"><span data-stu-id="474ea-121">Run the following command and then answer the prompts as follows.</span></span>

    ```
    yo office
    ```

    * <span data-ttu-id="474ea-122">Choisissez un type de projet : `Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="474ea-122">Choose a project type: `Excel Custom Functions Add-in project (...)`</span></span>

    * <span data-ttu-id="474ea-123">Choisissez un type de script : `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="474ea-123">Choose a script type: `JavaScript`</span></span>

    * <span data-ttu-id="474ea-124">Comment souhaitez-vous nommer votre complément ?</span><span class="sxs-lookup"><span data-stu-id="474ea-124">What do you want to name your add-in?</span></span> `stock-ticker`

    ![Le générateur de yeoman pour les compléments Office vous invite pour les fonctions personnalisées](../images/12-10-fork-cf-pic.jpg)

    <span data-ttu-id="474ea-126">Le générateur crée le projet et installe les composants Node.js de la prise en charge.</span><span class="sxs-lookup"><span data-stu-id="474ea-126">The generator will create the project and install supporting Node components.</span></span> <span data-ttu-id="474ea-127">Les fichiers de projet proviennent des référentiels [fonctions personnalisées Excel](https://github.com/OfficeDev/Excel-Custom-Functions)GitHub.</span><span class="sxs-lookup"><span data-stu-id="474ea-127">The project files come from the [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub repository.</span></span>

2. <span data-ttu-id="474ea-128">Accédez au dossier du projet.</span><span class="sxs-lookup"><span data-stu-id="474ea-128">Go to the project folder.</span></span>

    ```
    cd stock-ticker
    ```

3. <span data-ttu-id="474ea-129">Approuver le certificat auto-signé est nécessaire pour exécuter ce projet.</span><span class="sxs-lookup"><span data-stu-id="474ea-129">Trust the self-signed certificate that is needed to run this project.</span></span> <span data-ttu-id="474ea-130">Pour obtenir des instructions détaillées pour Windows ou Mac, voir [Ajout des Certificats Auto-signés comme Certificat Racine Approuvé](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span><span class="sxs-lookup"><span data-stu-id="474ea-130">For detailed instructions for either Windows or Mac, see [Adding Self Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span>  

4. <span data-ttu-id="474ea-131">Construire le projet.</span><span class="sxs-lookup"><span data-stu-id="474ea-131">Build the project.</span></span>

    ```
    npm run build
    ```

5. <span data-ttu-id="474ea-132">Démarrez le serveur web local qui est exécuté dans Node.js.</span><span class="sxs-lookup"><span data-stu-id="474ea-132">Start the local web server, which runs in Node.js.</span></span>

    * <span data-ttu-id="474ea-133">Si vous utilisez Excel pour Windows pour tester vos fonctions personnalisées, exécutez la commande suivante pour démarrer le serveur web local, ouvrir Excel et charger le complément :</span><span class="sxs-lookup"><span data-stu-id="474ea-133">If you'll be using Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:</span></span>

        ```
         npm run start
        ```
        <span data-ttu-id="474ea-134">Après avoir exécuté cette commande, votre invite de commandes affiche les détails sur ce que vous avez terminé, une autre fenêtre npm s’ouvre et affiche les détails de la génération et Excel commence par votre complément chargé.</span><span class="sxs-lookup"><span data-stu-id="474ea-134">After running this command, your command prompt will show details about what has been done, another npm window will open showing the details of the build, and Excel will start with your add-in loaded.</span></span> <span data-ttu-id="474ea-135">Si vous complément ne charge pas, vérifiez que vous avez correctement terminé l’étape 3.</span><span class="sxs-lookup"><span data-stu-id="474ea-135">If you add-in does not load, check that you have completed step 3 properly.</span></span>

    * <span data-ttu-id="474ea-136">Si vous utilisez Excel Online pour tester vos fonctions personnalisées, exécutez la commande suivante pour démarrer le serveur web local :</span><span class="sxs-lookup"><span data-stu-id="474ea-136">If you'll be using Excel Online to test your custom functions, run the following command to start the local web server:</span></span>

        ```
        npm run start-web
        ```

         <span data-ttu-id="474ea-137">Après avoir exécuté cette commande, une autre fenêtre s’ouvre et affiche les détails de la génération.</span><span class="sxs-lookup"><span data-stu-id="474ea-137">After running this command, another window will open showing you the details of the build.</span></span> <span data-ttu-id="474ea-138">Pour utiliser les fonctions, ouvrez un nouveau classeur dans Office Online.</span><span class="sxs-lookup"><span data-stu-id="474ea-138">To use your functions, open a new workbook in Office Online.</span></span>

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="474ea-139">Essayer une fonction personnalisée prédéfinie</span><span class="sxs-lookup"><span data-stu-id="474ea-139">Try out a prebuilt custom function</span></span>

<span data-ttu-id="474ea-140">Le projet de fonctions personnalisées que vous avez créé à l’aide du générateur Office Yo contient certaines fonctions personnalisées prédéfinies, définies dans le fichier **src/customfunction.js**.</span><span class="sxs-lookup"><span data-stu-id="474ea-140">The custom functions project that you created by using the Yo Office generator contains some prebuilt custom functions, defined within the **src/functions/functions.js** file.</span></span> <span data-ttu-id="474ea-141">Le fichier**manifest.xml**dans le répertoire racine du projet indique que toutes les fonctions personnalisées appartiennent à l’ `CONTOSO` espace de noms.</span><span class="sxs-lookup"><span data-stu-id="474ea-141">The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="474ea-142">Essayez de reproduire la`ADD` fonction personnalisée en complétant les étapes suivantes dans un classeur Excel :</span><span class="sxs-lookup"><span data-stu-id="474ea-142">In your Excel workbook, try out the `ADD` custom function by completing the following steps in Excel:</span></span>

1. <span data-ttu-id="474ea-143">Dans une cellule, tapez **= CONTOSO**.</span><span class="sxs-lookup"><span data-stu-id="474ea-143">Within a cell, type **=CONTOSO**.</span></span> <span data-ttu-id="474ea-144">Notez que le menu de saisie semi-automatique affiche la liste de toutes les fonctions dans l’`CONTOSO` espace de noms.</span><span class="sxs-lookup"><span data-stu-id="474ea-144">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="474ea-145">Exécutez la`CONTOSO.ADD` fonction, avec les nombres `10` et `200` comme paramètres d’entrée, en spécifiant la valeur`=CONTOSO.ADD(10,200)`suivante dans la cellule et appuyez sur entrée.</span><span class="sxs-lookup"><span data-stu-id="474ea-145">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="474ea-146">Le `ADD` fonction personnalisée calcule la somme des deux nombres que vous avez spécifiés comme paramètres d’entrée.</span><span class="sxs-lookup"><span data-stu-id="474ea-146">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="474ea-147">La saisie de`=CONTOSO.ADD(10,200)` doit générer le résultat **210** dans la cellule une fois que vous appuyez sur ENTRÉE.</span><span class="sxs-lookup"><span data-stu-id="474ea-147">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="474ea-148">Créer une fonction personnalisée qui demande les données à partir du web</span><span class="sxs-lookup"><span data-stu-id="474ea-148">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="474ea-149">Que se passe-t-il si vous avez besoin d’une fonction qui peut demander le prix d’une action à partir d’une API et afficher le résultat dans la cellule d’une feuille de calcul ?</span><span class="sxs-lookup"><span data-stu-id="474ea-149">What if you needed a function that could request the price of a stock from an API and display the result in the cell of a worksheet?</span></span> <span data-ttu-id="474ea-150">Les fonctions personnalisées sont conçues de sorte que vous pouvez facilement demander les données à partir du web de façon asynchrone.</span><span class="sxs-lookup"><span data-stu-id="474ea-150">Custom functions are designed so that you can easily request data from the web asynchronously.</span></span>

<span data-ttu-id="474ea-151">Procédez comme suit pour créer une fonction personnalisée nommée `stockPrice` qui accepte une action (par exemple, **MSFT**) et renvoie le prix de cette action.</span><span class="sxs-lookup"><span data-stu-id="474ea-151">Complete the following steps to create a custom function named `stockPrice` that accepts a stock ticker (e.g., **MSFT**) and returns the price of that stock.</span></span> <span data-ttu-id="474ea-152">Cette fonction personnalisée utilise l’API de cotation IEX, qui est gratuit et ne requiert pas d’authentification.</span><span class="sxs-lookup"><span data-stu-id="474ea-152">This custom function uses the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="474ea-153">Dans le projet **Bourse** que le Générateur de Yo Office a créé, recherchez le fichier**src/customfunctions.js** et ouvrez-le dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="474ea-153">In the **stock-ticker** project that the Yo Office generator created, find the file **src/customfunctions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="474ea-154">Dans**customfunctions.js**, recherchez la`increment` fonction et ajoutez le code suivant immédiatement après cette fonction.</span><span class="sxs-lookup"><span data-stu-id="474ea-154">In **customfunctions.js**, locate the `increment` function and add the following code immediately after that function.</span></span>

    ```js
    function stockPrice(ticker) {
        var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
        return fetch(url)
            .then(function(response) {
                return response.text();
            })
            .then(function(text) {
                return parseFloat(text);
            });

        // Note: in case of an error, the returned rejected Promise
        //    will be bubbled up to Excel to indicate an error.
    }

3. In **customfunctions.js**, locate the line`CustomFunctionMappings.INCREMENT = increment;`, add the following line of code immediately after that line, and save the file.

    ```js
    CustomFunctionMappings.STOCKPRICE = stockPrice;
    ```

4. <span data-ttu-id="474ea-155">Avant qu’Excel puisse rendre cette nouvelle fonction disponible, vous devez spécifier les métadonnées qui décrivent cette fonction à Excel.</span><span class="sxs-lookup"><span data-stu-id="474ea-155">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="474ea-156">Ouvrez le fichier**config/customfunctions.json**.</span><span class="sxs-lookup"><span data-stu-id="474ea-156">Open the **config/customfunctions.json** file.</span></span> <span data-ttu-id="474ea-157">Ajoutez l’objet JSON suivante à la matrice « fonctions » et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="474ea-157">Add the following object to the  array within the src/functions/functions.json file and save the file.</span></span>

    <span data-ttu-id="474ea-158">Cet élément JSON décrit la`stockPrice` fonction.</span><span class="sxs-lookup"><span data-stu-id="474ea-158">This JSON describes the `stockPrice` function.</span></span>

    ```JSON
    {
        "id": "STOCKPRICE",
        "name": "STOCKPRICE",
        "description": "Fetches current stock price",
        "helpUrl": "http://www.contoso.com/help",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock ticker name",
                "type": "string",
                "dimensionality": "scalar"
            }
        ]
    }
    ```

5. <span data-ttu-id="474ea-159">Vous devez réenregistrer le complément dans Excel afin que la nouvelle fonction soit disponible pour les utilisateurs finaux.</span><span class="sxs-lookup"><span data-stu-id="474ea-159">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="474ea-160">Complétez les étapes pour la plateforme que vous utiliserez dorénavant dans ce didacticiel.</span><span class="sxs-lookup"><span data-stu-id="474ea-160">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="474ea-161">Si vous utilisez Excel pour Windows :</span><span class="sxs-lookup"><span data-stu-id="474ea-161">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="474ea-162">Fermez Excel, puis ouvrez de nouveau Excel.</span><span class="sxs-lookup"><span data-stu-id="474ea-162">Close Excel and then reopen Excel.</span></span>

        2. <span data-ttu-id="474ea-163">Dans Excel, sélectionnez l’onglet**insérer**, puis cliquez sur la flèche vers le bas située à droite de **mes compléments**.  ![Insérer du ruban dans Excel pour Windows avec la flèche Mes compléments mise en évidence](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="474ea-163">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        3. <span data-ttu-id="474ea-164">Dans la liste des compléments disponibles, recherchez la section **Compléments Développeur** et sélectionnez votre complément**bourse** pour effectuer cette opération.</span><span class="sxs-lookup"><span data-stu-id="474ea-164">In the list of available add-ins, find the Developer Add-ins section and select the your add-in to register it.</span></span>
            <span data-ttu-id="474ea-165">![Insérer du ruban dans Excel pour Windows avec le complément Fonctions personnalisées Excel mis en évidence dans la liste Mes compléments](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="474ea-165">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="474ea-166">Si vous utilisez Excel Online :</span><span class="sxs-lookup"><span data-stu-id="474ea-166">If you're using Excel Online:</span></span>

        1. <span data-ttu-id="474ea-167">Dans Excel Online, sélectionnez l’onglet **insérer**, puis **compléments**. ![Insérer du ruban dans Excel Online avec l’icône Mes compléments mis en évidence](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="474ea-167">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="474ea-168">Sélectionnez**Gérer mes compléments** et sélectionnez **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="474ea-168">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="474ea-169">Sélectionnez \*\*Parcourir... \*\* et accédez au répertoire racine du projet créé par le Générateur de Yo Office.</span><span class="sxs-lookup"><span data-stu-id="474ea-169">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="474ea-170">Sélectionnez le fichier**manifest.xml** puis sélectionnez**Ouvrir**, puis sélectionnez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="474ea-170">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

6. <span data-ttu-id="474ea-171">À présent, nous allons essayer la nouvelle fonction.</span><span class="sxs-lookup"><span data-stu-id="474ea-171">Now, let's try out the new function.</span></span> <span data-ttu-id="474ea-172">Dans la cellule **B1**, tapez le texte `=CONTOSO.STOCKPRICE("MSFT")` et appuyez sur ENTRÉE.</span><span class="sxs-lookup"><span data-stu-id="474ea-172">In cell **B1**, type the text `=CONTOSO.STOCKPRICE("MSFT")` and press enter.</span></span> <span data-ttu-id="474ea-173">Vous devriez voir que le résultat dans la cellule **B1** est le prix boursier actuel pour un partage de stock Microsoft.</span><span class="sxs-lookup"><span data-stu-id="474ea-173">You should see that the result in cell **B1** is the current stock price for one share of Microsoft stock.</span></span>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="474ea-174">Créer une fonction personnalisée asynchrone diffusion en continu</span><span class="sxs-lookup"><span data-stu-id="474ea-174">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="474ea-175">La `stockPrice` fonction que vous venez de créer renvoie le prix d’une action à un moment donné, mais les prix des actions changent constamment.</span><span class="sxs-lookup"><span data-stu-id="474ea-175">The `stockPrice` function that you just created returns the price of a stock at a specific moment in time, but stock prices are always changing.</span></span> <span data-ttu-id="474ea-176">Nous allons créer une fonction personnalisée des flux de données à partir d’une API pour obtenir des mises à jour en temps réel sur un prix boursier.</span><span class="sxs-lookup"><span data-stu-id="474ea-176">Let's create a custom function that streams data from an API to get real-time updates on a stock price.</span></span>

<span data-ttu-id="474ea-177">Procédez comme suit pour créer une fonction personnalisée nommée `stockPriceStream` qui demande le prix d’une action boursière spécifique chaque 1000 millisecondes (à condition que la demande précédente soit terminée).</span><span class="sxs-lookup"><span data-stu-id="474ea-177">Complete the following steps to create a custom function named `stockPriceStream` that requests the price of the specified stock every 1000 milliseconds (provided that the previous request has completed).</span></span> <span data-ttu-id="474ea-178">Pendant la requête initiale en cours, vous pourrez afficher la valeur de l’espace réservé **## CHARGEMENT_DONNEES** la cellule dans laquelle la fonction est appelée.</span><span class="sxs-lookup"><span data-stu-id="474ea-178">While the initial request is in-progress, you may see the placeholder value **#GETTING_DATA** the cell where the function is being called.</span></span> <span data-ttu-id="474ea-179">Lorsqu’une valeur est renvoyée par la fonction **## CHARGEMENT_DONNEES** sera remplacée par cette valeur dans la cellule.</span><span class="sxs-lookup"><span data-stu-id="474ea-179">When a value is returned by the function, **#GETTING_DATA** will be replaced by that value in the cell.</span></span>

1. <span data-ttu-id="474ea-180">Dans le projet **Bourse** que le Générateur de Yo Office a créé, ajoutez le fichier **src/customfunctions.js** et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="474ea-180">In the **stock-ticker** project that the Yo Office generator created, add the following code to **src/customfunctions.js** and save the file.</span></span>

    ```js
    function stockPriceStream(ticker, handler) {
        var updateFrequency = 1000 /* milliseconds*/;
        var isPending = false;

        var timer = setInterval(function() {
            // If there is already a pending request, skip this iteration:
            if (isPending) {
                return;
            }

            var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
            isPending = true;

            fetch(url)
                .then(function(response) {
                    return response.text();
                })
                .then(function(text) {
                    handler.setResult(parseFloat(text));
                })
                .catch(function(error) {
                    handler.setResult(error);
                })
                .then(function() {
                    isPending = false;
                });
        }, updateFrequency);

        handler.onCanceled = () => {
            clearInterval(timer);
        };
    }

    CustomFunctionMappings.STOCKPRICESTREAM = stockPriceStream;
    ```

2. <span data-ttu-id="474ea-181">Avant qu’Excel puisse rendre cette nouvelle fonction disponible aux utilisateurs, vous devez spécifier les métadonnées qui décrivent cette fonction.</span><span class="sxs-lookup"><span data-stu-id="474ea-181">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="474ea-182">Dans le projet **Bourse** que le Générateur de Yo Office a créé, ajoutez l’objet suivant à la `functions` matrice au sein du fichier**config/customfunctions.json** et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="474ea-182">In the **stock-ticker** project that the Yo Office generator created, add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="474ea-183">Cet élément JSON décrit la`stockPriceStream` fonction.</span><span class="sxs-lookup"><span data-stu-id="474ea-183">This JSON describes the `stockPriceStream` function.</span></span> <span data-ttu-id="474ea-184">Pour n’importe quelle fonction de diffusion en continu, la propriété`stream` et la propriété`cancelable`doivent être définies `true` au sein de l’ `options` objet, comme illustré dans cet exemple de code.</span><span class="sxs-lookup"><span data-stu-id="474ea-184">For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.</span></span>

    ```json
    { 
        "id": "STOCKPRICESTREAM",
        "name": "STOCKPRICESTREAM",
        "description": "Streams real time stock price",
        "helpUrl": "http://www.contoso.com/help",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock ticker name",
                "type": "string",
                "dimensionality": "scalar"
            }
        ],
        "options": {
            "stream": true,
            "cancelable": true
        }
    }
    ```

3. <span data-ttu-id="474ea-185">Vous devez réenregistrer le complément dans Excel afin que la nouvelle fonction soit disponible pour les utilisateurs finaux.</span><span class="sxs-lookup"><span data-stu-id="474ea-185">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="474ea-186">Complétez les étapes pour la plateforme que vous utiliserez dorénavant dans ce didacticiel.</span><span class="sxs-lookup"><span data-stu-id="474ea-186">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="474ea-187">Si vous utilisez Excel pour Windows :</span><span class="sxs-lookup"><span data-stu-id="474ea-187">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="474ea-188">Fermez Excel, puis ouvrez de nouveau Excel.</span><span class="sxs-lookup"><span data-stu-id="474ea-188">Close Excel and then reopen Excel.</span></span>
        
        2. <span data-ttu-id="474ea-189">Dans Excel, sélectionnez l’onglet**insérer**, puis cliquez sur la flèche vers le bas située à droite de **mes compléments**.  ![Insérer du ruban dans Excel pour Windows avec la flèche Mes compléments mise en évidence](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="474ea-189">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        3. <span data-ttu-id="474ea-190">Dans la liste des compléments disponibles, recherchez la section **Compléments Développeur** et sélectionnez votre complément**bourse** pour effectuer cette opération.</span><span class="sxs-lookup"><span data-stu-id="474ea-190">In the list of available add-ins, find the Developer Add-ins section and select the your add-in to register it.</span></span>
            <span data-ttu-id="474ea-191">![Insérer du ruban dans Excel pour Windows avec le complément Fonctions personnalisées Excel mis en évidence dans la liste Mes compléments](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="474ea-191">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="474ea-192">Si vous utilisez Excel Online :</span><span class="sxs-lookup"><span data-stu-id="474ea-192">If you're using Excel Online:</span></span>

        1. <span data-ttu-id="474ea-193">Dans Excel Online, sélectionnez l’onglet **insérer**, puis **compléments**. ![Insérer du ruban dans Excel Online avec l’icône Mes compléments mis en évidence](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="474ea-193">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="474ea-194">Sélectionnez**Gérer mes compléments** et sélectionnez **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="474ea-194">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

        3. <span data-ttu-id="474ea-195">Sélectionnez \*\*Parcourir... \*\* et accédez au répertoire racine du projet créé par le Générateur de Yo Office.</span><span class="sxs-lookup"><span data-stu-id="474ea-195">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span>

        4. <span data-ttu-id="474ea-196">Sélectionnez le fichier**manifest.xml** puis sélectionnez**Ouvrir**, puis sélectionnez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="474ea-196">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

4. <span data-ttu-id="474ea-197">À présent, nous allons essayer la nouvelle fonction.</span><span class="sxs-lookup"><span data-stu-id="474ea-197">Now, let's try out the new function.</span></span> <span data-ttu-id="474ea-198">Dans la cellule **C1**, tapez le texte `=CONTOSO.STOCKPRICESTREAM("MSFT")` et appuyez sur ENTRÉE.</span><span class="sxs-lookup"><span data-stu-id="474ea-198">In cell **C1**, type the text `=CONTOSO.STOCKPRICESTREAM("MSFT")` and press enter.</span></span> <span data-ttu-id="474ea-199">Si le marché est ouvert, vous devriez voir que le résultat dans la cellule **C1** constamment mis à jour pour refléter le prix en temps réel pour un partage d’actions Microsoft.</span><span class="sxs-lookup"><span data-stu-id="474ea-199">Provided that the stock market is open, you should see that the result in cell **C1** is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span>

## <a name="next-steps"></a><span data-ttu-id="474ea-200">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="474ea-200">Next steps</span></span>

<span data-ttu-id="474ea-201">Dans ce didacticiel, vous avez créé un nouveau projet de fonctions personnalisées, essayé une fonction prédéfinie, créé une fonction personnalisée qui demande les données à partir du web et créé une fonction personnalisée qui diffuse les données en temps réel à partir du web.</span><span class="sxs-lookup"><span data-stu-id="474ea-201">In this tutorial, you've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web.</span></span> <span data-ttu-id="474ea-202">Pour en savoir plus sur les fonctions personnalisées dans Excel, passez à l’article suivant :</span><span class="sxs-lookup"><span data-stu-id="474ea-202">To learn more about custom functions in Excel, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="474ea-203">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="474ea-203">Create custom functions in Excel</span></span>](../excel/custom-functions-overview.md)

## <a name="legal-information"></a><span data-ttu-id="474ea-204">Informations légales</span><span class="sxs-lookup"><span data-stu-id="474ea-204">Legal information</span></span>

<span data-ttu-id="474ea-205">Données fournies gratuitement par [IEX](https://iextrading.com/developer/).</span><span class="sxs-lookup"><span data-stu-id="474ea-205">Data provided free by [IEX](https://iextrading.com/developer/).</span></span> <span data-ttu-id="474ea-206">Afficher les [conditions d’utilisation de IEX](https://iextrading.com/api-exhibit-a/).</span><span class="sxs-lookup"><span data-stu-id="474ea-206">View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/).</span></span> <span data-ttu-id="474ea-207">L’utilisation de Microsoft de l’API IEX dans ce didacticiel est uniquement à des fins d’enseignement.</span><span class="sxs-lookup"><span data-stu-id="474ea-207">Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>
