# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="b47b8-101">Didacticiel : créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="b47b8-101">Create custom functions in Excel (Preview)</span></span>

## <a name="introduction"></a><span data-ttu-id="b47b8-102">Présentation</span><span class="sxs-lookup"><span data-stu-id="b47b8-102">Introduction</span></span>

<span data-ttu-id="b47b8-103">Les fonctions personnalisées vous permettent d’ajouter de nouvelles fonctions dans Excel en définissant ces fonctions dans JavaScript dans le cadre d’un complément.</span><span class="sxs-lookup"><span data-stu-id="b47b8-103">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="b47b8-104">Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe quelle fonction native d’Excel, telle que `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="b47b8-104">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="b47b8-105">Vous pouvez créer des fonctions personnalisées qui effectuent des tâches simples telles que des calculs personnalisés ou des tâches plus complexes telles que la diffusion en continu des données en temps réel à partir du web dans une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="b47b8-105">You can create custom functions that perform simple tasks such as custom calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="b47b8-106">Dans ce didacticiel, vous allez :</span><span class="sxs-lookup"><span data-stu-id="b47b8-106">In this tutorial you will load AngularJS from CDN.</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="b47b8-107">Créer un projet de fonctions personnalisées à l’aide du générateur Yo Office</span><span class="sxs-lookup"><span data-stu-id="b47b8-107">Create a custom functions project by using the Yo Office generator</span></span>
> * <span data-ttu-id="b47b8-108">Utiliser une fonction personnalisée prédéfinie pour effectuer un calcul simple</span><span class="sxs-lookup"><span data-stu-id="b47b8-108">Use a prebuilt custom function to perform a simple calculation</span></span>
> * <span data-ttu-id="b47b8-109">Créer une fonction personnalisée qui demande les données à partir du web</span><span class="sxs-lookup"><span data-stu-id="b47b8-109">Create a custom function that requests data from the web</span></span>
> * <span data-ttu-id="b47b8-110">Créer une fonction personnalisée qui diffuse les données en temps réel à partir du web</span><span class="sxs-lookup"><span data-stu-id="b47b8-110">Create a custom function that streams real-time data from the web</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a><span data-ttu-id="b47b8-111">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="b47b8-111">Prerequisites</span></span>

* [<span data-ttu-id="b47b8-112">Node.js et npm</span><span class="sxs-lookup"><span data-stu-id="b47b8-112">Node and npm</span></span>](https://nodejs.org/en/)

* <span data-ttu-id="b47b8-113">[Git Bash](https://git-scm.com/downloads) (ou un autre client Git)</span><span class="sxs-lookup"><span data-stu-id="b47b8-113">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

* <span data-ttu-id="b47b8-114">La dernière version de [Yeoman](http://yeoman.io/) et le [générateur Yo Office](https://www.npmjs.com/package/generator-office).</span><span class="sxs-lookup"><span data-stu-id="b47b8-114">The latest version of [Yeoman](http://yeoman.io/) and the [Yo Office generator](https://www.npmjs.com/package/generator-office).</span></span> <span data-ttu-id="b47b8-115">À l’invite de commandes, exécutez la commande suivante pour installer ces outils :</span><span class="sxs-lookup"><span data-stu-id="b47b8-115">To install these tools globally, run the following command via the command prompt:</span></span>

    ```bash
    npm install -g yo generator-office
    ```

* <span data-ttu-id="b47b8-116">Excel pour Windows (1810 ou version ultérieure) ou Excel Online</span><span class="sxs-lookup"><span data-stu-id="b47b8-116">Excel for Windows (version 1810 or later) or Excel Online</span></span>

* <span data-ttu-id="b47b8-117">Rejoignez le[programme Office Insider](https://products.office.com/office-insider)(\*\* Insider\*\*niveau, anciennement appelé « Insider Fast »)</span><span class="sxs-lookup"><span data-stu-id="b47b8-117">Join the [Office Insider program](https://products.office.com/office-insider) (**Insider** level -- formerly called "Insider Fast")</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="b47b8-118">Créer un projet de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="b47b8-118">Create a custom functions project</span></span>

<span data-ttu-id="b47b8-119">Ce didacticiel commence à l’aide du générateur Yo Office pour créer les fichiers dont vous avez besoin pour votre projet fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="b47b8-119">You’ll begin this tutorial by using the Yo Office generator to create the files that you need for your custom functions project.</span></span>

1. <span data-ttu-id="b47b8-120">Exécutez la commande suivante, puis répondez aux invites comme suit.</span><span class="sxs-lookup"><span data-stu-id="b47b8-120">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    * <span data-ttu-id="b47b8-121">Choisissez un type de projet : `Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="b47b8-121">Choose a project type  </span></span>
    * <span data-ttu-id="b47b8-122">Choisissez un type de script : `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="b47b8-122">Choose a script type  </span></span>
    * <span data-ttu-id="b47b8-123">Comment souhaitez-vous nommer votre complément ?</span><span class="sxs-lookup"><span data-stu-id="b47b8-123">What do you want to name your add-in?</span></span> `stock-ticker`

    ![Yo Office bash vous invite pour fonctions personnalisées](../images/yo-office-cfs-stock-ticker-3.png)

    <span data-ttu-id="b47b8-125">Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants Node de la prise en charge.</span><span class="sxs-lookup"><span data-stu-id="b47b8-125">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span> <span data-ttu-id="b47b8-126">Les fichiers de projet proviennent des référentiels [fonctions personnalisées Excel](https://github.com/OfficeDev/Excel-Custom-Functions)GitHub.</span><span class="sxs-lookup"><span data-stu-id="b47b8-126">The project files come from the [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub repository.</span></span>

2. <span data-ttu-id="b47b8-127">Accédez au dossier du projet.</span><span class="sxs-lookup"><span data-stu-id="b47b8-127">Navigate to the project folder:</span></span>

    ```bash
    cd stock-ticker
    ```

3. <span data-ttu-id="b47b8-128">Démarrez le serveur web local.</span><span class="sxs-lookup"><span data-stu-id="b47b8-128">Start the local web server.</span></span>

    * <span data-ttu-id="b47b8-129">Si vous utilisez Excel pour Windows pour tester vos fonctions personnalisées, exécutez la commande suivante pour démarrer le serveur web local, ouvrir Excel et charger le complément :</span><span class="sxs-lookup"><span data-stu-id="b47b8-129">If you'll be using Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:</span></span>

        ```bash
        npm start
        ```

    * <span data-ttu-id="b47b8-130">Si vous utilisez Excel pour Windows pour tester vos fonctions personnalisées, exécutez la commande suivante pour démarrer le serveur web local :</span><span class="sxs-lookup"><span data-stu-id="b47b8-130">If you'll be using Excel Online to test your custom functions, run the following command to start the local web server:</span></span> 

        ```bash
        npm run start-web
        ```

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="b47b8-131">Essayer une fonction personnalisée prédéfinie</span><span class="sxs-lookup"><span data-stu-id="b47b8-131">Try out a prebuilt custom function</span></span>

<span data-ttu-id="b47b8-132">Le projet de fonctions personnalisées que vous avez créé à l’aide du générateur Office Yo contient certaines fonctions personnalisées prédéfinies, définies dans le **src/customfunction.js** fichier.</span><span class="sxs-lookup"><span data-stu-id="b47b8-132">The custom functions project that you created by using the Yo Office generator contains some prebuilt custom functions, defined within the **src/customfunction.js** file.</span></span> <span data-ttu-id="b47b8-133">Le **manifest.xml** fichier dans le répertoire racine du projet indique que toutes les fonctions personnalisées appartiennent à l’ `CONTOSO` espace de noms.</span><span class="sxs-lookup"><span data-stu-id="b47b8-133">The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="b47b8-134">Avant de pouvoir utiliser les fonctions personnalisées prédéfinies, vous devez inscrire le complément fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="b47b8-134">Before you can use any of the prebuilt custom functions, you must register the custom functions add-in in Excel.</span></span> <span data-ttu-id="b47b8-135">Pour cela, complétez les étapes pour la plateforme que vous utiliserez dorénavant dans ce didacticiel.</span><span class="sxs-lookup"><span data-stu-id="b47b8-135">Do so by completing steps for the platform that you'll be using in this tutorial.</span></span>

* <span data-ttu-id="b47b8-136">Si vous utilisez Excel pour Windows pour tester vos fonctions personnalisées :</span><span class="sxs-lookup"><span data-stu-id="b47b8-136">If you'll be using Excel for Windows to test your custom functions:</span></span>

    1. <span data-ttu-id="b47b8-137">Dans Excel, sélectionnez l’onglet**insérer**, puis cliquez sur la flèche vers le bas située à droite de **mes compléments**.  ![Insérer du ruban dans Excel pour Windows avec la flèche Mes compléments mise en évidence](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="b47b8-137">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

    2. <span data-ttu-id="b47b8-138">Dans la liste des compléments disponibles, recherchez la section **Compléments développeur** et sélectionnez le complément **Fonctions personnalisées Excel** pour l’enregistrer.</span><span class="sxs-lookup"><span data-stu-id="b47b8-138">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
        <span data-ttu-id="b47b8-139">![Insérer du ruban dans Excel pour Windows avec le complément Fonctions personnalisées Excel mis en évidence dans la liste Mes compléments](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="b47b8-139">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

* <span data-ttu-id="b47b8-140">Si vous utilisez Excel Online pour tester vos fonctions personnalisées :</span><span class="sxs-lookup"><span data-stu-id="b47b8-140">If you'll be using Excel Online to test your custom functions:</span></span> 

    1. <span data-ttu-id="b47b8-141">Dans Excel Online, sélectionnez l’onglet **insérer**, puis **compléments**. ![Insérer du ruban dans Excel Online avec l’icône Mes compléments mis en évidence](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="b47b8-141">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

    2. <span data-ttu-id="b47b8-142">Sélectionnez **Gérer mes compléments** et sélectionnez **Charger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="b47b8-142">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

    3. <span data-ttu-id="b47b8-143">Sélectionnez \*\*Parcourir... \*\* et accédez au répertoire racine du projet créé par le Générateur de Yo Office.</span><span class="sxs-lookup"><span data-stu-id="b47b8-143">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

    4. <span data-ttu-id="b47b8-144">Sélectionnez le fichier **manifest.xml** puis choisissez**Ouvrir**, puis sélectionnez **Charger**.</span><span class="sxs-lookup"><span data-stu-id="b47b8-144">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

<span data-ttu-id="b47b8-145">À ce stade, les fonctions personnalisées prédéfinies dans votre projet sont chargées et disponibles dans Excel.</span><span class="sxs-lookup"><span data-stu-id="b47b8-145">At this point, the prebuilt custom functions in your project are loaded and available within Excel.</span></span> <span data-ttu-id="b47b8-146">Essayez de reproduire la`ADD` fonction personnalisée en complétant les étapes suivantes dans Excel :</span><span class="sxs-lookup"><span data-stu-id="b47b8-146">Try out the `ADD` custom function by completing the following steps in Excel:</span></span>

1. <span data-ttu-id="b47b8-147">Dans une cellule, tapez **= CONTOSO**.</span><span class="sxs-lookup"><span data-stu-id="b47b8-147">Within a cell, type **=CONTOSO**.</span></span> <span data-ttu-id="b47b8-148">Notez que le menu de saisie semi-automatique affiche la liste de toutes les fonctions dans l’`CONTOSO` espace de noms.</span><span class="sxs-lookup"><span data-stu-id="b47b8-148">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="b47b8-149">Exécutez la `CONTOSO.ADD` fonction, avec les nombres `10` et `200` comme paramètres d’entrée, en spécifiant la valeur suivante dans la cellule et appuyez sur entrée :</span><span class="sxs-lookup"><span data-stu-id="b47b8-149">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by specifying the following value in the cell and pressing enter:</span></span>

    ```
    =CONTOSO.ADD(10,200)
    ```

<span data-ttu-id="b47b8-150">Le `ADD` fonction personnalisée calcule la somme des deux nombres que vous avez spécifiés comme paramètres d’entrée.</span><span class="sxs-lookup"><span data-stu-id="b47b8-150">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="b47b8-151">La saisie de`=CONTOSO.ADD(10,200)` doit générer le résultat **210** dans la cellule une fois que vous appuyez sur ENTRÉE.</span><span class="sxs-lookup"><span data-stu-id="b47b8-151">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="b47b8-152">Créer une fonction personnalisée qui demande les données à partir du web</span><span class="sxs-lookup"><span data-stu-id="b47b8-152">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="b47b8-153">Que se passe-t-il si vous avez besoin d’une fonction qui peut demander le prix d’une action à partir d’une API et afficher le résultat dans la cellule d’une feuille de calcul ?</span><span class="sxs-lookup"><span data-stu-id="b47b8-153">What if you needed a function that could request the price of a stock from an API and display the result in the cell of a worksheet?</span></span> <span data-ttu-id="b47b8-154">Les fonctions personnalisées sont conçues de sorte que vous pouvez facilement demander les données à partir du web de façon asynchrone.</span><span class="sxs-lookup"><span data-stu-id="b47b8-154">Custom functions are designed so that you can easily request data from the web asynchronously.</span></span>

<span data-ttu-id="b47b8-155">Procédez comme suit pour créer une fonction personnalisée nommée `stockPrice` qui accepte une action (par exemple, **MSFT**) et renvoie le prix de cette action.</span><span class="sxs-lookup"><span data-stu-id="b47b8-155">Complete the following steps to create a custom function named `stockPrice` that accepts a stock ticker (e.g., **MSFT**) and returns the price of that stock.</span></span> <span data-ttu-id="b47b8-156">Cette fonction personnalisée utilise l’API de cotation IEX, qui est gratuit et ne requiert pas d’authentification.</span><span class="sxs-lookup"><span data-stu-id="b47b8-156">This custom function uses the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="b47b8-157">Dans le projet **Bourse** que le Générateur de Yo Office a créé, recherchez le fichier **src/customfunctions.js** et ouvrez-le dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="b47b8-157">In the **stock-ticker** project that the Yo Office generator created, find the file **src/customfunctions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="b47b8-158">Ajoutez le code suivant à **customfunctions.js** et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="b47b8-158">Add the following code to **home.js** and save the file.</span></span>

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

    CustomFunctionMappings.STOCKPRICE = stockPrice;
    ```

3. <span data-ttu-id="b47b8-159">Avant qu’Excel puisse rendre cette nouvelle fonction disponible aux utilisateurs finaux, vous devez spécifier les métadonnées qui décrivent cette fonction.</span><span class="sxs-lookup"><span data-stu-id="b47b8-159">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="b47b8-160">Dans le projet **Bourse** que le Générateur de Yo Office a créé, recherchez le fichier **config/customfunctions.json** et ouvrez-le dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="b47b8-160">In the **stock-ticker** project that the Yo Office generator created, find the file **config/customfunctions.json** and open it in your code editor.</span></span> <span data-ttu-id="b47b8-161">Ajouter l’objet suivant à la `functions` matrice au sein du fichier **config/customfunctions.json** et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="b47b8-161">Add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="b47b8-162">JSON décrit la `stockPrice` fonction.</span><span class="sxs-lookup"><span data-stu-id="b47b8-162">This JSON describes the `stockPrice` function.</span></span>

    ```json
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

4. <span data-ttu-id="b47b8-163">Vous devez réenregistrer le complément dans Excel afin que la nouvelle fonction soit disponible pour les utilisateurs finaux.</span><span class="sxs-lookup"><span data-stu-id="b47b8-163">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="b47b8-164">Complétez les étapes pour la plateforme que vous utiliserez dorénavant dans ce didacticiel.</span><span class="sxs-lookup"><span data-stu-id="b47b8-164">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="b47b8-165">Si vous utilisez Excel pour Windows :</span><span class="sxs-lookup"><span data-stu-id="b47b8-165">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="b47b8-166">Fermez Excel, puis ouvrez de nouveau Excel.</span><span class="sxs-lookup"><span data-stu-id="b47b8-166">Close Excel and then reopen Excel.</span></span>

        2. <span data-ttu-id="b47b8-167">Dans Excel, sélectionnez l’onglet**insérer**, puis cliquez sur la flèche vers le bas située à droite de **mes compléments**.  ![Insérer du ruban dans Excel pour Windows avec la flèche Mes compléments mise en évidence](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="b47b8-167">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        1. <span data-ttu-id="b47b8-168">Dans la liste des compléments disponibles, recherchez la section **Compléments développeur** et sélectionnez le complément **Fonctions personnalisées Excel** pour l’enregistrer.</span><span class="sxs-lookup"><span data-stu-id="b47b8-168">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
            <span data-ttu-id="b47b8-169">![Insérer du ruban dans Excel pour Windows avec le complément Fonctions personnalisées Excel mis en évidence dans la liste Mes compléments](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="b47b8-169">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="b47b8-170">Si vous utilisez Excel Online :</span><span class="sxs-lookup"><span data-stu-id="b47b8-170">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="b47b8-171">Dans Excel Online, sélectionnez l’onglet **insérer**, puis **compléments**. ![Insérer du ruban dans Excel Online avec l’icône Mes compléments mis en évidence](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="b47b8-171">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="b47b8-172">Sélectionnez **Gérer mes compléments** et sélectionnez **Charger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="b47b8-172">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="b47b8-173">Sélectionnez \*\*Parcourir... \*\* et accédez au répertoire racine du projet créé par le Générateur de Yo Office.</span><span class="sxs-lookup"><span data-stu-id="b47b8-173">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="b47b8-174">Sélectionnez le fichier **manifest.xml** puis choisissez**Ouvrir**, puis sélectionnez **Charger**.</span><span class="sxs-lookup"><span data-stu-id="b47b8-174">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

5. <span data-ttu-id="b47b8-175">À présent, nous allons essayer la nouvelle fonction.</span><span class="sxs-lookup"><span data-stu-id="b47b8-175">Now, let's try out the new function.</span></span> <span data-ttu-id="b47b8-176">Dans la cellule **B1**, tapez le texte `=CONTOSO.STOCKPRICE("MSFT")` et appuyez sur ENTRÉE.</span><span class="sxs-lookup"><span data-stu-id="b47b8-176">In cell **B1**, type the text `=CONTOSO.STOCKPRICE("MSFT")` and press enter.</span></span> <span data-ttu-id="b47b8-177">Vous devriez voir que le résultat dans la cellule **B1** est le prix boursier actuel pour un partage de stock Microsoft.</span><span class="sxs-lookup"><span data-stu-id="b47b8-177">You should see that the result in cell **B1** is the current stock price for one share of Microsoft stock.</span></span>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="b47b8-178">Créer une fonction personnalisée asynchrone diffusion en continu</span><span class="sxs-lookup"><span data-stu-id="b47b8-178">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="b47b8-179">La `stockPrice` fonction que vous venez de créer renvoie le prix d’une action à un moment donné, mais les prix des actions changent constamment.</span><span class="sxs-lookup"><span data-stu-id="b47b8-179">The `stockPrice` function that you just created returns the price of a stock at a specific moment in time, but stock prices are always changing.</span></span> <span data-ttu-id="b47b8-180">Nous allons créer une fonction personnalisée des flux de données à partir d’une API pour obtenir des mises à jour en temps réel sur un prix boursier.</span><span class="sxs-lookup"><span data-stu-id="b47b8-180">Let's create a custom function that streams data from an API to get real-time updates on a stock price.</span></span>

<span data-ttu-id="b47b8-181">Procédez comme suit pour créer une fonction personnalisée nommée `stockPriceStream` qui demande le prix d’une action boursière spécifique chaque 1000 millisecondes (à condition que la demande précédente soit terminée).</span><span class="sxs-lookup"><span data-stu-id="b47b8-181">Complete the following steps to create a custom function named `stockPriceStream` that requests the price of the specified stock every 1000 milliseconds (provided that the previous request has completed).</span></span> <span data-ttu-id="b47b8-182">Pendant la requête initiale en cours, vous pourrez afficher la valeur de l’espace réservé **## CHARGEMENT_DONNEES** la cellule dans laquelle la fonction est appelée.</span><span class="sxs-lookup"><span data-stu-id="b47b8-182">While the initial request is in-progress, you may see the placeholder value **#GETTING_DATA** the cell where the function is being called.</span></span> <span data-ttu-id="b47b8-183">Lorsqu’une valeur est renvoyée par la fonction **## CHARGEMENT_DONNEES** sera remplacée par cette valeur dans la cellule.</span><span class="sxs-lookup"><span data-stu-id="b47b8-183">When a value is returned by the function, **#GETTING_DATA** will be replaced by that value in the cell.</span></span>

1. <span data-ttu-id="b47b8-184">Dans le projet **Bourse** que le Générateur de Yo Office a créé, ajoutez le fichier **src/customfunctions.js** et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="b47b8-184">In the **stock-ticker** project that the Yo Office generator created, add the following code to **src/customfunctions.js** and save the file.</span></span>

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

2. <span data-ttu-id="b47b8-185">Avant qu’Excel puisse rendre cette nouvelle fonction disponible aux utilisateurs finaux, vous devez spécifier les métadonnées qui décrivent cette fonction.</span><span class="sxs-lookup"><span data-stu-id="b47b8-185">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="b47b8-186">Dans le projet **Bourse** que le Générateur de Yo Office a créé, ajoutez l’objet suivant à la `functions` matrice au sein du fichier**config/customfunctions.json** et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="b47b8-186">In the **stock-ticker** project that the Yo Office generator created, add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="b47b8-187">JSON décrit la `stockPriceStream` fonction.</span><span class="sxs-lookup"><span data-stu-id="b47b8-187">This JSON describes the `stockPriceStream` function.</span></span> <span data-ttu-id="b47b8-188">Pour n’importe quelle fonction diffusion en continu, la `stream` propriété et la `cancelable` propriété doivent être définies `true` au sein de l’ `options` objet, comme illustré dans cet exemple de code.</span><span class="sxs-lookup"><span data-stu-id="b47b8-188">For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.</span></span>

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

3. <span data-ttu-id="b47b8-189">Vous devez réenregistrer le complément dans Excel afin que la nouvelle fonction soit disponible pour les utilisateurs finaux.</span><span class="sxs-lookup"><span data-stu-id="b47b8-189">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="b47b8-190">Complétez les étapes pour la plateforme que vous utiliserez dorénavant dans ce didacticiel.</span><span class="sxs-lookup"><span data-stu-id="b47b8-190">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="b47b8-191">Si vous utilisez Excel pour Windows :</span><span class="sxs-lookup"><span data-stu-id="b47b8-191">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="b47b8-192">Fermez Excel, puis ouvrez de nouveau Excel.</span><span class="sxs-lookup"><span data-stu-id="b47b8-192">Close Excel and then reopen Excel.</span></span>
        
        2. <span data-ttu-id="b47b8-193">Dans Excel, sélectionnez l’onglet**insérer**, puis cliquez sur la flèche vers le bas située à droite de **mes compléments**.  ![Insérer du ruban dans Excel pour Windows avec la flèche Mes compléments mise en évidence](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="b47b8-193">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        3. <span data-ttu-id="b47b8-194">Dans la liste des compléments disponibles, recherchez la section **Compléments développeur** et sélectionnez le complément **Fonctions personnalisées Excel** pour l’enregistrer.</span><span class="sxs-lookup"><span data-stu-id="b47b8-194">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
            <span data-ttu-id="b47b8-195">![Insérer du ruban dans Excel pour Windows avec le complément Fonctions personnalisées Excel mis en évidence dans la liste Mes compléments](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="b47b8-195">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="b47b8-196">Si vous utilisez Excel Online :</span><span class="sxs-lookup"><span data-stu-id="b47b8-196">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="b47b8-197">Dans Excel Online, sélectionnez l’onglet **insérer**, puis **compléments**. ![Insérer du ruban dans Excel Online avec l’icône Mes compléments mis en évidence](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="b47b8-197">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="b47b8-198">Sélectionnez **Gérer mes compléments** et sélectionnez **Charger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="b47b8-198">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="b47b8-199">Sélectionnez \*\*Parcourir... \*\* et accédez au répertoire racine du projet créé par le Générateur de Yo Office.</span><span class="sxs-lookup"><span data-stu-id="b47b8-199">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="b47b8-200">Sélectionnez le fichier **manifest.xml** puis choisissez**Ouvrir**, puis sélectionnez **Charger**.</span><span class="sxs-lookup"><span data-stu-id="b47b8-200">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

4. <span data-ttu-id="b47b8-201">À présent, nous allons essayer la nouvelle fonction.</span><span class="sxs-lookup"><span data-stu-id="b47b8-201">Now, let's try out the new function.</span></span> <span data-ttu-id="b47b8-202">Dans la cellule **C1**, tapez le texte `=CONTOSO.STOCKPRICESTREAM("MSFT")` et appuyez sur ENTRÉE.</span><span class="sxs-lookup"><span data-stu-id="b47b8-202">In cell **C1**, type the text `=CONTOSO.STOCKPRICESTREAM("MSFT")` and press enter.</span></span> <span data-ttu-id="b47b8-203">Si le marché est ouvert, vous devriez voir que le résultat dans la cellule **C1** constamment mis à jour pour refléter le prix en temps réel pour un partage d’actions Microsoft.</span><span class="sxs-lookup"><span data-stu-id="b47b8-203">Provided that the stock market is open, you should see that the result in cell **C1** is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span>

## <a name="next-steps"></a><span data-ttu-id="b47b8-204">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="b47b8-204">Next steps</span></span>

<span data-ttu-id="b47b8-205">Dans ce didacticiel, vous avez créé un nouveau projet de fonctions personnalisées, essayé une fonction prédéfinie, créé une fonction personnalisée qui demande les données à partir du web et créé une fonction personnalisée qui diffuse les données en temps réel à partir du web.</span><span class="sxs-lookup"><span data-stu-id="b47b8-205">In this tutorial, you've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web.</span></span> <span data-ttu-id="b47b8-206">Pour en savoir plus sur les fonctions personnalisées dans Excel, passez à l’article suivant :</span><span class="sxs-lookup"><span data-stu-id="b47b8-206">To learn more about custom functions in Excel, continue to the following article:</span></span> 

> [!div class="nextstepaction"]
> [<span data-ttu-id="b47b8-207">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="b47b8-207">Create custom functions in Excel (Preview)</span></span>](../excel/custom-functions-overview.md)

## <a name="legal-information"></a><span data-ttu-id="b47b8-208">Informations légales</span><span class="sxs-lookup"><span data-stu-id="b47b8-208">Legal information</span></span>

<span data-ttu-id="b47b8-209">Données fournies gratuitement par [IEX](https://iextrading.com/developer/).</span><span class="sxs-lookup"><span data-stu-id="b47b8-209">Data provided free by [IEX](https://iextrading.com/developer/).</span></span> <span data-ttu-id="b47b8-210">Afficher les [conditions d’utilisation de IEX](https://iextrading.com/api-exhibit-a/).</span><span class="sxs-lookup"><span data-stu-id="b47b8-210">View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/).</span></span> <span data-ttu-id="b47b8-211">L’utilisation de Microsoft de l’API IEX dans ce didacticiel est uniquement à des fins d’enseignement.</span><span class="sxs-lookup"><span data-stu-id="b47b8-211">Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>
