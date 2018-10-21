# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="8dd69-101">Tutoriel : Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="8dd69-101">Create custom functions in Excel (Preview)</span></span>

## <a name="introduction"></a><span data-ttu-id="8dd69-102">Introduction</span><span class="sxs-lookup"><span data-stu-id="8dd69-102">Introduction</span></span>

<span data-ttu-id="8dd69-p101">Les fonctions personnalisées vous permettent d’ajouter de nouvelles fonctions à Excel en définissant ces fonctions dans JavaScript comme partie d’un complément. Les utilisateurs dans Excel peuvent accéder aux fonctions personnalisées de la même façon qu’une fonction native dans Excel, telle que `SUM()`. Vous pouvez créer des fonctions personnalisées qui effectuent des tâches simples comme des calculs personnalisés ou des tâches plus complexes, comme la diffusion en continu des données en temps réel à partir du site Web dans une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="8dd69-p101">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`. You can create custom functions that perform simple tasks such as custom calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="8dd69-106">Dans ce tutoriel, vous allez :</span><span class="sxs-lookup"><span data-stu-id="8dd69-106">In this tutorial, you will use Visual Studio Code.</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="8dd69-107">Créer un projet de fonctions personnalisées à l’aide du Générateur de Yo Office</span><span class="sxs-lookup"><span data-stu-id="8dd69-107">Create a custom functions project by using the Yo Office generator</span></span>
> * <span data-ttu-id="8dd69-108">Utiliser une fonction personnalisée prédéfinie pour effectuer un calcul simple</span><span class="sxs-lookup"><span data-stu-id="8dd69-108">Use a prebuilt custom function to perform a simple calculation</span></span>
> * <span data-ttu-id="8dd69-109">Créer une fonction personnalisée qui demande des données à partir du web</span><span class="sxs-lookup"><span data-stu-id="8dd69-109">Create a custom function that requests data from the web</span></span>
> * <span data-ttu-id="8dd69-110">Créer une fonction personnalisée qui transmet les données en temps réel à partir du web</span><span class="sxs-lookup"><span data-stu-id="8dd69-110">Create a custom function that streams real-time data from the web</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a><span data-ttu-id="8dd69-111">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="8dd69-111">Prerequisites</span></span>

* [<span data-ttu-id="8dd69-112">Node.js et npm</span><span class="sxs-lookup"><span data-stu-id="8dd69-112">Node and npm</span></span>](https://nodejs.org/en/)

* <span data-ttu-id="8dd69-113">[Git Bash](https://git-scm.com/downloads) (ou un autre client Git)</span><span class="sxs-lookup"><span data-stu-id="8dd69-113">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

* <span data-ttu-id="8dd69-p102">La dernière version de [Yeoman](http://yeoman.io/) et le [Générateur de Yo Office](https://www.npmjs.com/package/generator-office). Pour installer ces outils globalement, exécutez la commande suivante par le biais de l’invite de commandes :</span><span class="sxs-lookup"><span data-stu-id="8dd69-p102">The latest version of [Yeoman](http://yeoman.io/) and the [Yo Office generator](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```bash
    npm install -g yo generator-office
    ```

* <span data-ttu-id="8dd69-116">Excel pour Windows (numéro de build 10827 ou version ultérieure) ou Excel Online</span><span class="sxs-lookup"><span data-stu-id="8dd69-116">Excel for Windows (build number 10827 or later) or Excel Online</span></span>

* <span data-ttu-id="8dd69-117">Participer au [programme Office Insider](https://products.office.com/office-insider) (niveau**Insider** -- anciennement appelé « Insider Fast »)</span><span class="sxs-lookup"><span data-stu-id="8dd69-117">Join the [Office Insider program](https://products.office.com/office-insider) (**Insider** level -- formerly called "Insider Fast")</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="8dd69-118">Créer un projet de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="8dd69-118">Create a custom functions project by using the Yo Office generator</span></span>

<span data-ttu-id="8dd69-119">Vous allez commencer ce tutoriel à l’aide du Générateur de Yo Office pour créer les fichiers dont vous avez besoin pour votre projet de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="8dd69-119">You’ll begin this tutorial by using the Yo Office generator to create the files that you need for your custom functions project.</span></span>

1. <span data-ttu-id="8dd69-120">Exécutez la commande suivante, puis répondez aux invites comme suit.</span><span class="sxs-lookup"><span data-stu-id="8dd69-120">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    * <span data-ttu-id="8dd69-121">Choisissez un type de projet : `Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="8dd69-121">Choose a project type:`Excel Custom Functions Add-in project (...)`</span></span>
    * <span data-ttu-id="8dd69-122">Choisissez un type de script : `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="8dd69-122">Choose a script type:`JavaScript`</span></span>
    * <span data-ttu-id="8dd69-123">Comment souhaitez-vous nommer votre complément ?</span><span class="sxs-lookup"><span data-stu-id="8dd69-123">What do you want to name your add-in?:</span></span> `stock-ticker`

    ![L'invite bash yo office vous demande de  fournir des fonctions personnalisées](../images/yo-office-cfs-stock-ticker-3.png)

    <span data-ttu-id="8dd69-p103">Après avoir exécuté l’assistant, le générateur crée les fichiers du projet et installe les composants de nœud de la prise en charge. Les fichiers de projet viennent du référentiel [GitHub Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions).</span><span class="sxs-lookup"><span data-stu-id="8dd69-p103">After you complete the wizard, the generator will create the project files and install supporting Node components. The project files come from the [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub repository.</span></span>

2. <span data-ttu-id="8dd69-127">Accédez au dossier du projet.</span><span class="sxs-lookup"><span data-stu-id="8dd69-127">Navigate to the project folder:</span></span>

    ```bash
    cd stock-ticker
    ```

3. <span data-ttu-id="8dd69-128">Démarrez le serveur web local.</span><span class="sxs-lookup"><span data-stu-id="8dd69-128">Start the local web server.</span></span>

    * <span data-ttu-id="8dd69-129">Si vous utilisez Excel pour Windows pour tester vos fonctions personnalisées, exécutez la commande suivante pour démarrer le serveur web local, lancer Excel et charger en parallèle le complément :</span><span class="sxs-lookup"><span data-stu-id="8dd69-129">If you'll be using Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:</span></span>

        ```bash
        npm start
        ```

    * <span data-ttu-id="8dd69-130">Si vous allez utiliser Excel Online pour tester vos fonctions personnalisées, exécutez la commande suivante pour démarrer le serveur web local :</span><span class="sxs-lookup"><span data-stu-id="8dd69-130">If you'll be using Excel Online to test your custom functions, run the following command to start the local web server:</span></span> 

        ```bash
        npm run start-web
        ```

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="8dd69-131">Essayer une fonction personnalisée prédéfinie</span><span class="sxs-lookup"><span data-stu-id="8dd69-131">Try out a prebuilt custom function</span></span>

<span data-ttu-id="8dd69-p104">Le projet de fonctions personnalisées que vous avez créé à l’aide du Générateur de Yo Office contient certaines fonctions personnalisées prédéfinies au niveau du fichier **src/customfunction.js**. Le fichier **manifest.xml** dans le répertoire racine du projet spécifie que toutes les fonctions personnalisées appartiennent à l’espace de noms `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="8dd69-p104">The custom functions project that you created by using the Yo Office generator contains some prebuilt custom functions, defined within the **src/customfunction.js** file. The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="8dd69-p105">Avant de pouvoir utiliser une des fonctions personnalisées prédéfinies, vous devez enregistrer le complément fonctions personnalisées dans Excel. Faites cela en procédant comme pour la plateforme que vous utiliserez dans ce tutoriel.</span><span class="sxs-lookup"><span data-stu-id="8dd69-p105">Before you can use any of the prebuilt custom functions, you must register the custom functions add-in in Excel. Do so by completing steps for the platform that you'll be using in this tutorial.</span></span>

* <span data-ttu-id="8dd69-136">Si vous utilisez Excel pour Windows pour tester vos fonctions personnalisées :</span><span class="sxs-lookup"><span data-stu-id="8dd69-136">If you'll be using Excel for Windows to test your custom functions:</span></span>

    1. <span data-ttu-id="8dd69-137">Dans Excel, sélectionnez l’onglet **Insertion**, puis choisissez la flèche bas située à droite de **Mes applications**.  ![Insérez un ruban dans Excel pour Windows avec la flèche de Mes applications mise en surbrillance](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="8dd69-137">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

    2. <span data-ttu-id="8dd69-p106">Dans la liste des compléments disponibles, recherchez la section **Compléments pour développeurs** et sélectionnez le complément **Fonctions personnalisées d’Excel** pour l’enregistrer.  ![Insérez le ruban dans Excel pour Windows avec le complément des fonctions personnalisées d’Excel mis en surbrillance dans la liste Mes compléments.](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="8dd69-p106">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.  ![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

* <span data-ttu-id="8dd69-140">Si vous utilisez Excel Online pour tester vos fonctions personnalisées :</span><span class="sxs-lookup"><span data-stu-id="8dd69-140">If you'll be using Excel Online to test your custom functions:</span></span> 

    1. <span data-ttu-id="8dd69-141">Dans Excel Online, choisissez l’onglet **Insertion** , puis choisissez **Compléments**.  ![Insérez le ruban dans Excel Online avec l’icône Mes applications mise en surbrillance](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="8dd69-141">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

    2. <span data-ttu-id="8dd69-142">Sélectionnez **Gérer mes compléments** , sélectionnez **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="8dd69-142">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

    3. <span data-ttu-id="8dd69-143">Cliquez sur **Parcourir** et accédez au répertoire racine du projet que le Générateur de Yo Office a créé.</span><span class="sxs-lookup"><span data-stu-id="8dd69-143">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

    4. <span data-ttu-id="8dd69-144">Sélectionnez le fichier **manifest.xml** et choisissez **Ouvrir**, puis cliquez sur **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="8dd69-144">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

<span data-ttu-id="8dd69-p107">À ce stade, les fonctions personnalisées prédéfinies dans votre projet sont chargés et disponibles dans Excel. Essayez la fonction personnalisée `ADD` en effectuant les étapes suivantes dans Excel :</span><span class="sxs-lookup"><span data-stu-id="8dd69-p107">At this point, the prebuilt custom functions in your project are loaded and available within Excel. Try out the `ADD` custom function by completing the following steps in Excel:</span></span>

1. <span data-ttu-id="8dd69-p108">Dans une cellule, tapez **= CONTOSO**. Notez que le menu de saisie semi-automatique affiche la liste de toutes les fonctions dans l'espace de noms `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="8dd69-p108">Within a cell, type **=CONTOSO**. Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="8dd69-149">Exécutez la fonction `CONTOSO.ADD`, avec les numéros `10` et `200` comme paramètres d’entrée, en spécifiant la valeur suivante dans la cellule et en appuyant sur Entrée :</span><span class="sxs-lookup"><span data-stu-id="8dd69-149">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by specifying the following value in the cell and pressing enter:</span></span>

    ```
    =CONTOSO.ADD(10,200)
    ```

<span data-ttu-id="8dd69-p109">La fonction personnalisée `ADD` calcule la somme de deux nombres que vous spécifiez comme paramètres d’entrée. Si vous tapez `=CONTOSO.ADD(10,200)`, vous devez obtenir le résultat **210** dans la cellule lorsque vous appuyez sur Entrée.</span><span class="sxs-lookup"><span data-stu-id="8dd69-p109">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters. Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="8dd69-152">Créer une fonction personnalisée qui demande des données à partir du web</span><span class="sxs-lookup"><span data-stu-id="8dd69-152">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="8dd69-p110">Et si vous aviez besoin d’une fonction qui peut demander le prix d’une action à partir d’une API et afficher le résultat dans la cellule d’une feuille de calcul ? Les fonctions personnalisées sont conçues afin que vous puissiez aisément demander des données à partir du web de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="8dd69-p110">What if you needed a function that could request the price of a stock from an API and display the result in the cell of a worksheet? Custom functions are designed so that you can easily request data from the web asynchronously.</span></span>

<span data-ttu-id="8dd69-p111">Effectuez les étapes suivantes pour créer une fonction personnalisée nommée `stockPrice` qui a comme argument un symbole boursier (par exemple, **MSFT**) et renvoie le prix de l’action correspondante. Cette fonction personnalisée utilise l’API IEX Trading, qui est gratuite et ne nécessite pas d’authentification.</span><span class="sxs-lookup"><span data-stu-id="8dd69-p111">Complete the following steps to create a custom function named `stockPrice` that accepts a stock ticker (e.g., **MSFT**) and returns the price of that stock. This custom function uses the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="8dd69-157">Dans le projet de **symboles boursiers** créé par le générateur Yo Office, recherchez le fichier **src/customfunctions.js** et ouvrez-le dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="8dd69-157">In the **stock-ticker** project that the Yo Office generator created, find the file **src/customfunctions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="8dd69-158">Ajoutez le code suivant à **customfunctions.js** et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="8dd69-158">Add the following code to **home.js** and save the file.</span></span>

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

3. <span data-ttu-id="8dd69-p112">Avant qu’Excel ne puisse rendre cette nouvelle fonction disponible pour les utilisateurs finaux, vous devez spécifier les métadonnées décrivant cette fonction. Dans le projet de **symboles boursiers** créé par le générateur de Yo Office, recherchez le fichier **config/customfunctions.js** et ouvrez-le dans votre éditeur de code. Ajoutez l’objet suivant au tableau `functions` dans le fichier **config/customfunctions.json** et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="8dd69-p112">Before Excel can make this new function available to end-users, you must specify metadata that describes this function. In the **stock-ticker** project that the Yo Office generator created, find the file **config/customfunctions.json** and open it in your code editor. Add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="8dd69-162">Ce code JSON décrit la fonction `stockPrice`.</span><span class="sxs-lookup"><span data-stu-id="8dd69-162">This JSON describes the `stockPrice` function.</span></span>

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

4. <span data-ttu-id="8dd69-p113">Vous devez réenregistrer le complément dans Excel pour que la nouvelle fonction soit disponible pour les utilisateurs finaux. Effectuez les étapes suivantes pour la plateforme que vous utilisez dans ce tutoriel.</span><span class="sxs-lookup"><span data-stu-id="8dd69-p113">You must reregister the add-in in Excel in order for the new function to be available to end-users. Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="8dd69-165">Si vous utilisez Excel pour Windows :</span><span class="sxs-lookup"><span data-stu-id="8dd69-165">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="8dd69-166">Fermez Excel, puis rouvrez Excel.</span><span class="sxs-lookup"><span data-stu-id="8dd69-166">Close Excel and then reopen Excel.</span></span>

        2. <span data-ttu-id="8dd69-167">Dans Excel, sélectionnez l’onglet **Insertion**, puis choisissez la flèche bas située à droite de **Mes applications**.  ![Insérez un ruban dans Excel pour Windows avec la flèche de Mes applications mise en surbrillance](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="8dd69-167">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        1. <span data-ttu-id="8dd69-p114">Dans la liste des compléments disponibles, recherchez la section **Compléments pour développeurs** et sélectionnez le complément **Fonctions personnalisées d’Excel** pour l’enregistrer.  ![Insérez le ruban dans Excel pour Windows avec le complément des fonctions personnalisées d’Excel mis en surbrillance dans la liste Mes compléments.](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="8dd69-p114">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.  ![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="8dd69-170">Si vous utilisez Excel Online :</span><span class="sxs-lookup"><span data-stu-id="8dd69-170">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="8dd69-171">Dans Excel Online, choisissez l’onglet **Insertion** , puis choisissez **Compléments**.  ![Insérez le ruban dans Excel Online avec l’icône Mes applications mise en surbrillance](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="8dd69-171">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="8dd69-172">Sélectionnez **Gérer mes compléments** , sélectionnez **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="8dd69-172">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="8dd69-173">Cliquez sur **Parcourir** et accédez au répertoire racine du projet que le Générateur de Yo Office a créé.</span><span class="sxs-lookup"><span data-stu-id="8dd69-173">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="8dd69-174">Sélectionnez le fichier **manifest.xml** et choisissez **Ouvrir**, puis cliquez sur **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="8dd69-174">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

5. <span data-ttu-id="8dd69-p115">À présent, nous allons essayer la nouvelle fonction. Dans la cellule **B1**, tapez le texte `=CONTOSO.STOCKPRICE("MSFT")` et appuyez sur Entrée. Vous devriez voir que le résultat dans la cellule **B1** est le cours actuel d’une action Microsoft.</span><span class="sxs-lookup"><span data-stu-id="8dd69-p115">Now, let's try out the new function. In cell **B1**, type the text `=CONTOSO.STOCKPRICE("MSFT")` and press enter. You should see that the result in cell **B1** is the current stock price for one share of Microsoft stock.</span></span>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="8dd69-178">Créer une fonction personnalisée asynchrone en continu</span><span class="sxs-lookup"><span data-stu-id="8dd69-178">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="8dd69-p116">La fonction `stockPrice` que vous venez de créer renvoie le prix d’une action à un moment spécifique, mais les prix des actions varient constamment. Nous allons créer une fonction personnalisée qui récupère des flux de données à partir d’une API pour obtenir des mises à jour en temps réel sur le cours des actions.</span><span class="sxs-lookup"><span data-stu-id="8dd69-p116">The `stockPrice` function that you just created returns the price of a stock at a specific moment in time, but stock prices are always changing. Let's create a custom function that streams data from an API to get real-time updates on a stock price.</span></span>

<span data-ttu-id="8dd69-p117">Effectuez les étapes suivantes pour créer une fonction personnalisée nommée `stockPriceStream` qui demande le prix de l’action toutes les 1000 millisecondes (à condition que la requête précédente soit terminée). Pendant que la requête initiale est en cours, vous pouvez voir le message d’indication **## GETTING_DATA** au niveau de la cellule dans laquelle la fonction est appelée. Lorsqu’une valeur est retournée par la fonction, **#GETTING_DATA** sera remplacé par cette valeur dans la cellule.</span><span class="sxs-lookup"><span data-stu-id="8dd69-p117">Complete the following steps to create a custom function named `stockPriceStream` that requests the price of the specified stock every 1000 milliseconds (provided that the previous request has completed). While the initial request is in-progress, you may see the placeholder value **#GETTING_DATA** the cell where the function is being called. When a value is returned by the function, **#GETTING_DATA** will be replaced by that value in the cell.</span></span>

1. <span data-ttu-id="8dd69-184">Dans le projet **symboles boursiers** créé par le Générateur de Yo Office, ajoutez le code suivant à **src/customfunctions.js** et sauvegardez le fichier.</span><span class="sxs-lookup"><span data-stu-id="8dd69-184">In the **stock-ticker** project that the Yo Office generator created, add the following code to **src/customfunctions.js** and save the file.</span></span>

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

2. <span data-ttu-id="8dd69-p118">Avant qu’Excel ne puisse rendre cette nouvelle fonction disponible pour les utilisateurs finaux, vous devez spécifier les métadonnées décrivant cette fonction. Dans le projet de **symboles boursiers** créé par le générateur de Yo Office, ajoutez l’objet suivant au tableau `functions` dans le fichier **config/customfunctions.json** et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="8dd69-p118">Before Excel can make this new function available to end-users, you must specify metadata that describes this function. In the **stock-ticker** project that the Yo Office generator created, add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="8dd69-p119">Ce code JSON décrit la fonction `stockPriceStream`. Pour une fonction de diffusion en continu, les propriétés `stream` et `cancelable` doivent être définies à `true` dans l’objet `options`, comme illustré dans cet exemple de code.</span><span class="sxs-lookup"><span data-stu-id="8dd69-p119">This JSON describes the `stockPriceStream` function. For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.</span></span>

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

3. <span data-ttu-id="8dd69-p120">Vous devez réenregistrer le complément dans Excel pour que la nouvelle fonction soit disponible pour les utilisateurs finaux. Effectuez les étapes suivantes pour la plateforme que vous utilisez dans ce tutoriel.</span><span class="sxs-lookup"><span data-stu-id="8dd69-p120">You must reregister the add-in in Excel in order for the new function to be available to end-users. Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="8dd69-191">Si vous utilisez Excel pour Windows :</span><span class="sxs-lookup"><span data-stu-id="8dd69-191">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="8dd69-192">Fermez Excel, puis rouvrez Excel.</span><span class="sxs-lookup"><span data-stu-id="8dd69-192">Close Excel and then reopen Excel.</span></span>
        
        2. <span data-ttu-id="8dd69-193">Dans Excel, sélectionnez l’onglet **Insertion**, puis choisissez la flèche bas située à droite de **Mes applications**.  ![Insérez un ruban dans Excel pour Windows avec la flèche de Mes applications mise en surbrillance](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="8dd69-193">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        3. <span data-ttu-id="8dd69-p121">Dans la liste des compléments disponibles, recherchez la section **Compléments pour développeurs** et sélectionnez le complément **Fonctions personnalisées d’Excel** pour l’enregistrer.  ![Insérez le ruban dans Excel pour Windows avec le complément des fonctions personnalisées d’Excel mis en surbrillance dans la liste Mes compléments.](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="8dd69-p121">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.  ![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="8dd69-196">Si vous utilisez Excel Online :</span><span class="sxs-lookup"><span data-stu-id="8dd69-196">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="8dd69-197">Dans Excel Online, choisissez l’onglet **Insertion** , puis choisissez **Compléments**.  ![Insérez le ruban dans Excel Online avec l’icône Mes applications mise en surbrillance](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="8dd69-197">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="8dd69-198">Sélectionnez **Gérer mes compléments** , sélectionnez **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="8dd69-198">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="8dd69-199">Cliquez sur **Parcourir** et accédez au répertoire racine du projet que le Générateur de Yo Office a créé.</span><span class="sxs-lookup"><span data-stu-id="8dd69-199">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="8dd69-200">Sélectionnez le fichier **manifest.xml** et choisissez **Ouvrir**, puis cliquez sur **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="8dd69-200">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

4. <span data-ttu-id="8dd69-p122">À présent, nous allons essayer la nouvelle fonction. Dans la cellule **C1**, tapez le texte `=CONTOSO.STOCKPRICESTREAM("MSFT")` et appuyez sur Entrée. À condition que le marché boursier est ouvert, vous devez voir que le résultat dans la cellule **C1** est constamment mis à jour pour refléter le prix en temps réel pour une action Microsoft.</span><span class="sxs-lookup"><span data-stu-id="8dd69-p122">Now, let's try out the new function. In cell **C1**, type the text `=CONTOSO.STOCKPRICESTREAM("MSFT")` and press enter. Provided that the stock market is open, you should see that the result in cell **C1** is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span>

## <a name="next-steps"></a><span data-ttu-id="8dd69-204">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="8dd69-204">Next steps</span></span>

<span data-ttu-id="8dd69-p123">Dans ce tutoriel, vous avez créé un nouveau projet de fonctions personnalisées, essayé une fonction prédéfinie, créé une fonction personnalisée qui demande des données à partir du web et créé une fonction personnalisée qui récupère des flux de données en temps réel à partir du web. Pour en savoir plus sur les fonctions personnalisées dans Excel, passez à l’article suivant :</span><span class="sxs-lookup"><span data-stu-id="8dd69-p123">In this tutorial, you've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web. To learn more about custom functions in Excel, continue to the following article:</span></span> 

> [!div class="nextstepaction"]
> [<span data-ttu-id="8dd69-207">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="8dd69-207">Create custom functions in Excel (Preview)</span></span>](../excel/custom-functions-overview.md)

## <a name="legal-information"></a><span data-ttu-id="8dd69-208">Mentions légales</span><span class="sxs-lookup"><span data-stu-id="8dd69-208">Legal Information</span></span>

<span data-ttu-id="8dd69-p124">Données fournies gratuitement par [IEX](https://iextrading.com/developer/). Affichez les [conditions d’utilisation d’IEX](https://iextrading.com/api-exhibit-a/). L’utilisation de l’API IEX dans ce tutoriel par Microsoft est uniquement à des fins de formation.</span><span class="sxs-lookup"><span data-stu-id="8dd69-p124">Data provided free by [IEX](https://iextrading.com/developer/). View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/). Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>
