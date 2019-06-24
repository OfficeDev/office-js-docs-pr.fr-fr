---
title: Didacticiel de fonctions personnalisées Excel
description: Dans ce didacticiel, vous allez créer un complément Excel qui contient une fonction personnalisée qui effectue des calculs, requiert des données web ou lance un flux de données web.
ms.date: 06/20/2019
ms.prod: excel
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: 3ae7896c082e7a1a45fb153dc69772f206a433de
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35126980"
---
# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="eb86a-103">Didacticiel : créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="eb86a-103">Tutorial: Create custom functions in Excel</span></span>

<span data-ttu-id="eb86a-104">Les fonctions personnalisées vous permettent d’ajouter de nouvelles fonctions dans Excel en définissant ces fonctions dans JavaScript dans le cadre d’un complément.</span><span class="sxs-lookup"><span data-stu-id="eb86a-104">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="eb86a-105">Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe quelle fonction native d’Excel, telle que `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="eb86a-105">Users within Excel can access custom functions as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="eb86a-106">Vous pouvez créer des fonctions personnalisées qui effectuent des tâches simples comme des calculs personnalisés ou des tâches plus complexes telles que la diffusion en continu des données en temps réel à partir du web dans une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="eb86a-106">You can create custom functions that perform simple tasks like calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="eb86a-107">Dans ce didacticiel, vous allez :</span><span class="sxs-lookup"><span data-stu-id="eb86a-107">In this tutorial, you will:</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="eb86a-108">Créer un complément de fonction personnalisée à l’aide la [Générateur Yeoman de compléments Office](https://www.npmjs.com/package/generator-office).</span><span class="sxs-lookup"><span data-stu-id="eb86a-108">Create a custom function add-in using the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office).</span></span> 
> * <span data-ttu-id="eb86a-109">Utiliser une fonction personnalisée prédéfinie pour effectuer un calcul simple.</span><span class="sxs-lookup"><span data-stu-id="eb86a-109">Use a prebuilt custom function to perform a simple calculation.</span></span>
> * <span data-ttu-id="eb86a-110">Créer une fonction personnalisée qui demande les données à partir du web.</span><span class="sxs-lookup"><span data-stu-id="eb86a-110">Create a custom function that gets data from the web.</span></span>
> * <span data-ttu-id="eb86a-111">Créer une fonction personnalisée qui diffuse les données en temps réel à partir du web.</span><span class="sxs-lookup"><span data-stu-id="eb86a-111">Create a custom function that streams real-time data from the web.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="eb86a-112">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="eb86a-112">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* <span data-ttu-id="eb86a-113">Excel sur Windows (version 1904 ou ultérieure, connexion à l’abonnement Office 365) ou sur le Web</span><span class="sxs-lookup"><span data-stu-id="eb86a-113">Excel on Windows (version 1904 or later, connected to Office 365 subscription) or on the web</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="eb86a-114">Créer un projet de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="eb86a-114">Create a custom functions project</span></span>

 <span data-ttu-id="eb86a-115">Pour commencer, vous devez créer le projet de code pour créer votre complément de fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="eb86a-115">To start, you'll create the code project to build your custom function add-in.</span></span> <span data-ttu-id="eb86a-116">Le [Générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office) configurera votre projet avec certaines fonctions personnalisées prédéfinies que vous pouvez tester. Si vous avez déjà exécuté le démarrage rapide des fonctions personnalisées et généré un projet, continuez à utiliser ce projet et passez à [cette étape](#create-a-custom-function-that-requests-data-from-the-web) .</span><span class="sxs-lookup"><span data-stu-id="eb86a-116">The [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) will set up your project with some prebuilt custom functions that you can try out. If you have already run the custom functions quick start and generated a project, continue to use that project and skip to [this step](#create-a-custom-function-that-requests-data-from-the-web) instead.</span></span>

1. <span data-ttu-id="eb86a-117">Exécutez la commande suivante, puis répondez aux invitations comme suit.</span><span class="sxs-lookup"><span data-stu-id="eb86a-117">Run the following command and then answer the prompts as follows.</span></span>
    
    ```command&nbsp;line
    yo office
    ```
    
    * <span data-ttu-id="eb86a-118">**Sélectionnez un type de projet :** `Excel Custom Functions Add-in project`</span><span class="sxs-lookup"><span data-stu-id="eb86a-118">**Choose a project type:** `Excel Custom Functions Add-in project`</span></span>
    * <span data-ttu-id="eb86a-119">**Sélectionnez un type de script :** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="eb86a-119">**Choose a script type:** `JavaScript`</span></span>
    * <span data-ttu-id="eb86a-120">**Comment souhaitez-vous nommer votre complément ?**</span><span class="sxs-lookup"><span data-stu-id="eb86a-120">**What do you want to name your add-in?**</span></span> `stock-ticker`

    ![Le générateur de yeoman pour les compléments Office vous invite pour les fonctions personnalisées](../images/UpdatedYoOfficePrompt.png)
    
    <span data-ttu-id="eb86a-122">Le générateur crée le projet et installe les composants Node.js de la prise en charge.</span><span class="sxs-lookup"><span data-stu-id="eb86a-122">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="eb86a-123">Accédez au dossier racine du projet.</span><span class="sxs-lookup"><span data-stu-id="eb86a-123">Navigate to the root folder of the project.</span></span>
    
    ```command&nbsp;line
    cd stock-ticker
    ```

3. <span data-ttu-id="eb86a-124">Créez le projet.</span><span class="sxs-lookup"><span data-stu-id="eb86a-124">Build the project.</span></span>
    
    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > <span data-ttu-id="eb86a-125">Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez.</span><span class="sxs-lookup"><span data-stu-id="eb86a-125">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="eb86a-126">Si vous êtes invité à installer un certificat après avoir exécuté `npm run build`, acceptez d’installer le certificat fourni par le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="eb86a-126">If you are prompted to install a certificate after you run `npm run build`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

4. <span data-ttu-id="eb86a-127">Démarrez le serveur web local qui est exécuté dans Node.js.</span><span class="sxs-lookup"><span data-stu-id="eb86a-127">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="eb86a-128">Vous pouvez essayer le complément de fonction personnalisée dans Excel sur le Web ou Windows.</span><span class="sxs-lookup"><span data-stu-id="eb86a-128">You can try out the custom function add-in in Excel on the web or Windows.</span></span>

# <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="eb86a-129">Excel sur Windows</span><span class="sxs-lookup"><span data-stu-id="eb86a-129">Excel on Windows</span></span>](#tab/excel-windows)

<span data-ttu-id="eb86a-130">Pour tester votre complément dans Excel sous Windows, exécutez la commande suivante.</span><span class="sxs-lookup"><span data-stu-id="eb86a-130">To test your add-in in Excel on Windows, run the following command.</span></span> <span data-ttu-id="eb86a-131">Lorsque vous exécutez cette commande, le serveur Web local démarre et Excel s’ouvre avec votre complément chargé.</span><span class="sxs-lookup"><span data-stu-id="eb86a-131">When you run this command, the local web server will start and Excel will open with your add-in loaded.</span></span>

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-on-the-webtabexcel-online"></a>[<span data-ttu-id="eb86a-132">Excel sur le Web</span><span class="sxs-lookup"><span data-stu-id="eb86a-132">Excel on the web</span></span>](#tab/excel-online)

<span data-ttu-id="eb86a-133">Pour tester votre complément dans Excel sur un navigateur, exécutez la commande suivante.</span><span class="sxs-lookup"><span data-stu-id="eb86a-133">To test your add-in in Excel on a browser, run the following command.</span></span> <span data-ttu-id="eb86a-134">Lorsque vous exécutez cette commande, le serveur web local démarre.</span><span class="sxs-lookup"><span data-stu-id="eb86a-134">When you run this command, the local web server will start.</span></span>

```command&nbsp;line
npm run start:web
```

<span data-ttu-id="eb86a-135">Pour utiliser votre complément de fonctions personnalisées, ouvrez un nouveau classeur dans Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="eb86a-135">To use your custom functions add-in, open a new workbook in Excel on the web.</span></span> <span data-ttu-id="eb86a-136">Dans ce classeur, effectuez les étapes suivantes pour chargement votre complément.</span><span class="sxs-lookup"><span data-stu-id="eb86a-136">In this workbook, complete the following steps to sideload your add-in.</span></span>

1. <span data-ttu-id="eb86a-137">Dans Excel, sélectionnez l’onglet **insertion** , puis **compléments**.</span><span class="sxs-lookup"><span data-stu-id="eb86a-137">In Excel, choose the **Insert** tab and then choose **Add-ins**.</span></span>

   ![Insérer un ruban dans Excel sur le Web avec l’icône mes compléments mise en surbrillance](../images/excel-cf-online-register-add-in-1.png)
   
2. <span data-ttu-id="eb86a-139">Sélectionnez**Gérer mes compléments** et sélectionnez **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="eb86a-139">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="eb86a-140">Sélectionnez \*\*Parcourir... \*\* et accédez au répertoire racine du projet créé par le Générateur de Yo Office.</span><span class="sxs-lookup"><span data-stu-id="eb86a-140">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="eb86a-141">Sélectionnez le fichier**manifest.xml** puis sélectionnez**Ouvrir**, puis sélectionnez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="eb86a-141">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 
    
## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="eb86a-142">Essayer une fonction personnalisée prédéfinie</span><span class="sxs-lookup"><span data-stu-id="eb86a-142">Try out a prebuilt custom function</span></span>

<span data-ttu-id="eb86a-143">Le projet de fonctions personnalisées que vous avez créé contient des fonctions personnalisées prédéfinies, définies dans le fichier **./SRC/Functions/functions.js** .</span><span class="sxs-lookup"><span data-stu-id="eb86a-143">The custom functions project that you created contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file.</span></span> <span data-ttu-id="eb86a-144">Le fichier**manifest.xml**indique que toutes les fonctions personnalisées appartiennent à l’`CONTOSO`espace de noms.</span><span class="sxs-lookup"><span data-stu-id="eb86a-144">The **./manifest.xml** file specifies that all custom functions belong to the `CONTOSO` namespace.</span></span> <span data-ttu-id="eb86a-145">L’espace de noms CONTOSO permet d’accéder aux fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="eb86a-145">You'll use the CONTOSO namespace to access the custom functions in Excel.</span></span>

<span data-ttu-id="eb86a-146">Essayez de reproduire la`ADD` fonction personnalisée en complétant les étapes suivantes dans Excel:</span><span class="sxs-lookup"><span data-stu-id="eb86a-146">Next you'll try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="eb86a-147">Dans Excel, accédez à n’importe quelle cellule et entrez `=CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="eb86a-147">In Excel, go to any cell and enter `=CONTOSO`.</span></span> <span data-ttu-id="eb86a-148">Notez que le menu de saisie semi-automatique affiche la liste de toutes les fonctions dans l’espace de noms `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="eb86a-148">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="eb86a-149">Exécutez la`CONTOSO.ADD` fonction, avec les nombres `10` et `200` comme paramètres d’entrée, en spécifiant la valeur`=CONTOSO.ADD(10,200)`suivante dans la cellule et appuyez sur entrée.</span><span class="sxs-lookup"><span data-stu-id="eb86a-149">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="eb86a-150">Le `ADD` fonction personnalisée calcule la somme des deux nombres que vous avez spécifiés et renvoie le résultat**210** .</span><span class="sxs-lookup"><span data-stu-id="eb86a-150">The `ADD` custom function computes the sum of the two numbers that you provided and returns the result of **210**.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="eb86a-151">Créer une fonction personnalisée qui demande les données à partir du web</span><span class="sxs-lookup"><span data-stu-id="eb86a-151">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="eb86a-152">Intégration de données à partir du Web est un excellent moyen pour étendre Excel via les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="eb86a-152">Integrating data from the Web is a great way to extend Excel through custom functions.</span></span> <span data-ttu-id="eb86a-153">Vous allez ensuite créer une fonction personnalisée nommée `stockPrice` qui obtient des actions à partir d’une API Web et renvoie le résultat à la cellule d’une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="eb86a-153">Next you’ll create a custom function named `stockPrice` that gets a stock quote from a Web API and returns the result to the cell of a worksheet.</span></span> 

> [!NOTE]
> <span data-ttu-id="eb86a-154">Le code suivant demande une cotation boursière à l’aide de l’API commerce IEX.</span><span class="sxs-lookup"><span data-stu-id="eb86a-154">The following code requests a stock quote using the IEX Trading API.</span></span> <span data-ttu-id="eb86a-155">Avant de pouvoir exécuter le code, vous devez [créer un compte gratuit avec le Cloud Iex](https://iexcloud.io/) afin que vous puissiez obtenir le jeton d’API requis dans la demande d’API.</span><span class="sxs-lookup"><span data-stu-id="eb86a-155">Before you can run the code, you'll need to [create a free account with IEX Cloud](https://iexcloud.io/) so that you can get the API token that's required in the API request.</span></span>  

1. <span data-ttu-id="eb86a-156">Dans le projet **boursier** , recherchez le fichier **./SRC/Functions/functions.js** et ouvrez-le dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="eb86a-156">In the **stock-ticker** project, find the file **./src/functions/functions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="eb86a-157">Dans **functions. js**, recherchez `increment` la fonction et ajoutez le code suivant après cette fonction.</span><span class="sxs-lookup"><span data-stu-id="eb86a-157">In **functions.js**, locate the `increment` function and add the following code after that function.</span></span>

    ```js
    /**
    * Fetches current stock price
    * @customfunction 
    * @param {string} ticker Stock symbol
    * @returns {number} The current stock price.
    */
    function stockPrice(ticker) {
        //Note: In the following line, replace <YOUR_TOKEN_HERE> with the API token that you've obtained through your IEX Cloud account.
        var url = "https://cloud.iexapis.com/stable/stock/" + ticker + "/quote/latestPrice?token=<YOUR_TOKEN_HERE>"
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
    CustomFunctions.associate("STOCKPRICE", stockPrice);
    ```

    <span data-ttu-id="eb86a-158">Le `CustomFunctions.associate` code associe le `id`de la fonction avec l’adresse de la fonction de `stockPrice` dans JavaScript afin qu’Excel peut appeler votre fonction.</span><span class="sxs-lookup"><span data-stu-id="eb86a-158">The `CustomFunctions.associate` code associates the `id` of the function with the function address of `stockPrice` in JavaScript so that Excel can call your function.</span></span>

3. <span data-ttu-id="eb86a-159">Exécutez la commande suivante pour regénérer le projet.</span><span class="sxs-lookup"><span data-stu-id="eb86a-159">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

4. <span data-ttu-id="eb86a-160">Procédez comme suit (pour Excel sur le Web ou Windows) pour réenregistrer le complément dans Excel.</span><span class="sxs-lookup"><span data-stu-id="eb86a-160">Complete the following steps (for either Excel on the web or Windows) to re-register the add-in in Excel.</span></span> <span data-ttu-id="eb86a-161">Vous devez effectuer ces étapes avant que la nouvelle fonction ne soit disponible.</span><span class="sxs-lookup"><span data-stu-id="eb86a-161">You must complete these steps before the new function will be available.</span></span>

# <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="eb86a-162">Excel sur Windows</span><span class="sxs-lookup"><span data-stu-id="eb86a-162">Excel on Windows</span></span>](#tab/excel-windows)

1. <span data-ttu-id="eb86a-163">Fermez Excel, puis ouvrez de nouveau Excel.</span><span class="sxs-lookup"><span data-stu-id="eb86a-163">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="eb86a-164">Dans Excel, sélectionnez l’onglet **Insérer** , puis cliquez sur la flèche vers le bas située à droite de **mes compléments**.  ![Insérer un ruban dans Excel sur Windows avec la flèche mes compléments mise en surbrillance](../images/select-insert.png)</span><span class="sxs-lookup"><span data-stu-id="eb86a-164">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel on Windows with the My Add-ins arrow highlighted](../images/select-insert.png)</span></span>

3. <span data-ttu-id="eb86a-165">Dans la liste des compléments disponibles, recherchez la section **Compléments Développeur** et sélectionnez votre complément**bourse** pour effectuer cette opération.</span><span class="sxs-lookup"><span data-stu-id="eb86a-165">In the list of available add-ins, find the **Developer Add-ins** section and select the **stock-ticker** add-in to register it.</span></span>
    <span data-ttu-id="eb86a-166">![Insérer un ruban dans Excel sur Windows avec le complément de fonctions personnalisées Excel mis en surbrillance dans la liste mes compléments](../images/list-stock-ticker-red.png)</span><span class="sxs-lookup"><span data-stu-id="eb86a-166">![Insert ribbon in Excel on Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/list-stock-ticker-red.png)</span></span>

# <a name="excel-on-the-webtabexcel-online"></a>[<span data-ttu-id="eb86a-167">Excel sur le Web</span><span class="sxs-lookup"><span data-stu-id="eb86a-167">Excel on the web</span></span>](#tab/excel-online)

1. <span data-ttu-id="eb86a-168">Dans Excel, sélectionnez l’onglet **insertion** , puis **compléments**.  ![Insérer un ruban dans Excel sur le Web avec l’icône mes compléments mise en surbrillance](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="eb86a-168">In Excel, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel on the web with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="eb86a-169">Sélectionnez**Gérer mes compléments** et sélectionnez **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="eb86a-169">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="eb86a-170">Sélectionnez \*\*Parcourir... \*\* et accédez au répertoire racine du projet créé par le Générateur de Yo Office.</span><span class="sxs-lookup"><span data-stu-id="eb86a-170">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="eb86a-171">Sélectionnez le fichier**manifest.xml** puis sélectionnez**Ouvrir**, puis sélectionnez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="eb86a-171">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

<ol start="5">
<li> <span data-ttu-id="eb86a-172">Essayez la nouvelle fonction.</span><span class="sxs-lookup"><span data-stu-id="eb86a-172">Try out the new function.</span></span> <span data-ttu-id="eb86a-173">Dans la cellule <strong>B1</strong>, tapez le texte <strong>= CONTOSO. STOCKPRICE("MSFT")</strong> et appuyez sur ENTRÉE.</span><span class="sxs-lookup"><span data-stu-id="eb86a-173">In cell <strong>B1</strong>, type the text <strong>=CONTOSO.STOCKPRICE("MSFT")</strong> and press enter.</span></span> <span data-ttu-id="eb86a-174">Vous devriez voir que le résultat dans la cellule <strong>B1</strong> est le prix boursier actuel pour un partage de stock Microsoft.</span><span class="sxs-lookup"><span data-stu-id="eb86a-174">You should see that the result in cell <strong>B1</strong> is the current stock price for one share of Microsoft stock.</span></span></li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="eb86a-175">Créer une fonction personnalisée asynchrone diffusion en continu</span><span class="sxs-lookup"><span data-stu-id="eb86a-175">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="eb86a-176">La fonction`stockPrice`que vous venez de créer renvoie le prix d’une action à un moment donné, mais les prix des actions changent constamment.</span><span class="sxs-lookup"><span data-stu-id="eb86a-176">The `stockPrice` function returns the price of a stock at a specific moment in time, but stock prices are always changing.</span></span> <span data-ttu-id="eb86a-177">Vous allez ensuite créer une fonction personnalisée nommée `stockPriceStream` qui obtient le prix d’une action chaque 1000 millisecondes.</span><span class="sxs-lookup"><span data-stu-id="eb86a-177">Next you’ll create a custom function named `stockPriceStream` that gets the price of a stock every 1000 milliseconds.</span></span>

1. <span data-ttu-id="eb86a-178">Dans le projet **boursier** , ajoutez le code suivant à **./SRC/Functions/functions.js** et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="eb86a-178">In the **stock-ticker** project, add the following code to **./src/functions/functions.js** and save the file.</span></span>

    ```js
    /**
    * Streams real time stock price
    * @customfunction 
    * @param {string} ticker Stock symbol
    * @param {CustomFunctions.StreamingInvocation<number>} invocation
    */
    function stockPriceStream(ticker, invocation) {
        var updateFrequency = 1000 /* milliseconds*/;
        var isPending = false;

        var timer = setInterval(function() {
            // If there is already a pending request, skip this iteration:
            if (isPending) {
                return;
            }

            //Note: In the following line, replace <YOUR_TOKEN_HERE> with the API token that you've obtained through your IEX Cloud account.
            var url = "https://cloud.iexapis.com/stable/stock/" + ticker + "/quote/latestPrice?token=<YOUR_TOKEN_HERE>"
            isPending = true;

            fetch(url)
                .then(function(response) {
                    return response.text();
                })
                .then(function(text) {
                    invocation.setResult(parseFloat(text));
                })
                .catch(function(error) {
                    invocation.setResult(error);
                })
                .then(function() {
                    isPending = false;
                });
        }, updateFrequency);

        invocation.onCanceled = () => {
            clearInterval(timer);
        };
    }
    CustomFunctions.associate("STOCKPRICESTREAM", stockPriceStream);
    ```
    
    <span data-ttu-id="eb86a-179">Le `CustomFunctions.associate` code associe le `id`de la fonction avec l’adresse de la fonction de `stockPriceStream` dans JavaScript afin qu’Excel peut appeler votre fonction.</span><span class="sxs-lookup"><span data-stu-id="eb86a-179">The `CustomFunctions.associate` code associates the `id` of the function with the function address of `stockPriceStream` in JavaScript so that Excel can call your function.</span></span>
    
2. <span data-ttu-id="eb86a-180">Exécutez la commande suivante pour regénérer le projet.</span><span class="sxs-lookup"><span data-stu-id="eb86a-180">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

3. <span data-ttu-id="eb86a-181">Procédez comme suit (pour Excel sur le Web ou Windows) pour réenregistrer le complément dans Excel.</span><span class="sxs-lookup"><span data-stu-id="eb86a-181">Complete the following steps (for either Excel on the web or Windows) to re-register the add-in in Excel.</span></span> <span data-ttu-id="eb86a-182">Vous devez effectuer ces étapes avant que la nouvelle fonction ne soit disponible.</span><span class="sxs-lookup"><span data-stu-id="eb86a-182">You must complete these steps before the new function will be available.</span></span> 

# <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="eb86a-183">Excel sur Windows</span><span class="sxs-lookup"><span data-stu-id="eb86a-183">Excel on Windows</span></span>](#tab/excel-windows)

1. <span data-ttu-id="eb86a-184">Fermez Excel, puis ouvrez de nouveau Excel.</span><span class="sxs-lookup"><span data-stu-id="eb86a-184">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="eb86a-185">Dans Excel, sélectionnez l’onglet **Insérer** , puis cliquez sur la flèche vers le bas située à droite de **mes compléments**.  ![Insérer un ruban dans Excel sur Windows avec la flèche mes compléments mise en surbrillance](../images/select-insert.png)</span><span class="sxs-lookup"><span data-stu-id="eb86a-185">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel on Windows with the My Add-ins arrow highlighted](../images/select-insert.png)</span></span>

3. <span data-ttu-id="eb86a-186">Dans la liste des compléments disponibles, recherchez la section **Compléments Développeur** et sélectionnez votre complément**bourse** pour effectuer cette opération.</span><span class="sxs-lookup"><span data-stu-id="eb86a-186">In the list of available add-ins, find the **Developer Add-ins** section and select the **stock-ticker** add-in to register it.</span></span>
    <span data-ttu-id="eb86a-187">![Insérer un ruban dans Excel sur Windows avec le complément de fonctions personnalisées Excel mis en surbrillance dans la liste mes compléments](../images/list-stock-ticker-red.png)</span><span class="sxs-lookup"><span data-stu-id="eb86a-187">![Insert ribbon in Excel on Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/list-stock-ticker-red.png)</span></span>

# <a name="excel-on-the-webtabexcel-online"></a>[<span data-ttu-id="eb86a-188">Excel sur le Web</span><span class="sxs-lookup"><span data-stu-id="eb86a-188">Excel on the web</span></span>](#tab/excel-online)

1. <span data-ttu-id="eb86a-189">Dans Excel, sélectionnez l’onglet **insertion** , puis **compléments**.  ![Insérer un ruban dans Excel sur le Web avec l’icône mes compléments mise en surbrillance](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="eb86a-189">In Excel, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel on the web with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="eb86a-190">Sélectionnez**Gérer mes compléments** et sélectionnez **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="eb86a-190">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="eb86a-191">Sélectionnez \*\*Parcourir... \*\* et accédez au répertoire racine du projet créé par le Générateur de Yo Office.</span><span class="sxs-lookup"><span data-stu-id="eb86a-191">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="eb86a-192">Sélectionnez le fichier**manifest.xml** puis sélectionnez**Ouvrir**, puis sélectionnez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="eb86a-192">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 

<ol start="4">
<li><span data-ttu-id="eb86a-193">Essayez la nouvelle fonction.</span><span class="sxs-lookup"><span data-stu-id="eb86a-193">Try out the new function.</span></span> <span data-ttu-id="eb86a-194">Dans la cellule <strong>C1</strong>, tapez le texte <strong>= CONTOSO. STOCKPRICE("MSFT")</strong> et appuyez sur ENTRÉE.</span><span class="sxs-lookup"><span data-stu-id="eb86a-194">In cell <strong>C1</strong>, type the text <strong>=CONTOSO.STOCKPRICESTREAM("MSFT")</strong> and press enter.</span></span> <span data-ttu-id="eb86a-195">Si le marché est ouvert, vous devriez voir que le résultat dans la cellule <strong>C1</strong> constamment mis à jour pour refléter le prix en temps réel pour un partage d’actions Microsoft.</span><span class="sxs-lookup"><span data-stu-id="eb86a-195">Provided that the stock market is open, you should see that the result in cell <strong>C1</strong> is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span></li>
</ol>

## <a name="next-steps"></a><span data-ttu-id="eb86a-196">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="eb86a-196">Next steps</span></span>

<span data-ttu-id="eb86a-197">Félicitations !</span><span class="sxs-lookup"><span data-stu-id="eb86a-197">Congratulations!</span></span> <span data-ttu-id="eb86a-198">Vous avez créé un nouveau projet de fonctions personnalisées, essayé une fonction prédéfinie, créé une fonction personnalisée qui demande les données à partir du web et créé une fonction personnalisée qui diffuse les données en temps réel à partir du web.</span><span class="sxs-lookup"><span data-stu-id="eb86a-198">You've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web.</span></span> <span data-ttu-id="eb86a-199">Vous pouvez également essayer de déboguer cette fonction à l’aide [des instructions de débogage de la fonction personnalisée](../excel/custom-functions-debugging.md).</span><span class="sxs-lookup"><span data-stu-id="eb86a-199">You can also try out debugging this function using [the custom function debugging instructions](../excel/custom-functions-debugging.md).</span></span> <span data-ttu-id="eb86a-200">Pour en savoir plus sur les fonctions personnalisées dans Excel, passez à l’article suivant :</span><span class="sxs-lookup"><span data-stu-id="eb86a-200">To learn more about custom functions in Excel, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="eb86a-201">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="eb86a-201">Create custom functions in Excel</span></span>](../excel/custom-functions-overview.md)

### <a name="legal-information"></a><span data-ttu-id="eb86a-202">Informations légales</span><span class="sxs-lookup"><span data-stu-id="eb86a-202">Legal information</span></span>

<span data-ttu-id="eb86a-203">Données fournies gratuitement par [IEX](https://iextrading.com/developer/).</span><span class="sxs-lookup"><span data-stu-id="eb86a-203">Data provided free by [IEX](https://iextrading.com/developer/).</span></span> <span data-ttu-id="eb86a-204">Afficher les [conditions d’utilisation de IEX](https://iextrading.com/api-exhibit-a/).</span><span class="sxs-lookup"><span data-stu-id="eb86a-204">View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/).</span></span> <span data-ttu-id="eb86a-205">L’utilisation de Microsoft de l’API IEX dans ce didacticiel est uniquement à des fins d’enseignement.</span><span class="sxs-lookup"><span data-stu-id="eb86a-205">Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>
