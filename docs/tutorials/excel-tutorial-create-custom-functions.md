---
title: Didacticiel de fonctions personnalisées Excel (aperçu)
description: Dans ce didacticiel, vous allez créer un complément Excel qui contient une fonction personnalisée qui effectue des calculs, requiert des données web ou lance un flux de données web.
ms.date: 01/08/2019
ms.topic: tutorial
ms.openlocfilehash: 46a9883e9dbc2e3bfbbe170665d82826bdfb26f9
ms.sourcegitcommit: 9afcb1bb295ec0c8940ed3a8364dbac08ef6b382
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/08/2019
ms.locfileid: "27770643"
---
# <a name="tutorial-create-custom-functions-in-excel-preview"></a><span data-ttu-id="bc031-103">Didacticiel : créer des fonctions personnalisées dans Excel (aperçu)</span><span class="sxs-lookup"><span data-stu-id="bc031-103">Tutorial: Create custom functions in Excel</span></span>

<span data-ttu-id="bc031-104">Les fonctions personnalisées vous permettent d’ajouter de nouvelles fonctions dans Excel en définissant ces fonctions dans JavaScript dans le cadre d’un complément.</span><span class="sxs-lookup"><span data-stu-id="bc031-104">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="bc031-105">Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe quelle fonction native d’Excel, telle que `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="bc031-105">Users within Excel can access custom functions as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="bc031-106">Vous pouvez créer des fonctions personnalisées qui effectuent des tâches simples comme des calculs personnalisés ou des tâches plus complexes telles que la diffusion en continu des données en temps réel à partir du web dans une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="bc031-106">You can create custom functions that perform simple tasks like calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="bc031-107">Dans ce didacticiel, vous allez :</span><span class="sxs-lookup"><span data-stu-id="bc031-107">In this tutorial, you will:</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="bc031-108">Créer un complément de fonction personnalisée à l’aide la [Générateur Yeoman de compléments Office](https://www.npmjs.com/package/generator-office).</span><span class="sxs-lookup"><span data-stu-id="bc031-108">Create a custom function add-in using the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office).</span></span> 
> * <span data-ttu-id="bc031-109">Utiliser une fonction personnalisée prédéfinie pour effectuer un calcul simple.</span><span class="sxs-lookup"><span data-stu-id="bc031-109">Use a prebuilt custom function to perform a simple calculation</span></span>
> * <span data-ttu-id="bc031-110">Créer une fonction personnalisée qui demande les données à partir du web.</span><span class="sxs-lookup"><span data-stu-id="bc031-110">Create a custom function that requests data from the web</span></span>
> * <span data-ttu-id="bc031-111">Créer une fonction personnalisée qui diffuse les données en temps réel à partir du web.</span><span class="sxs-lookup"><span data-stu-id="bc031-111">Create a custom function that streams real-time data from the web</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a><span data-ttu-id="bc031-112">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bc031-112">Prerequisites</span></span>

* <span data-ttu-id="bc031-113">[Node.js](https://nodejs.org/en/) (version 8.0.0 ou ultérieure)</span><span class="sxs-lookup"><span data-stu-id="bc031-113">[Node.js](https://nodejs.org/en/) (version 8.0.0 or later)</span></span>

* <span data-ttu-id="bc031-114">[Git Bash](https://git-scm.com/downloads) (ou un autre client Git)</span><span class="sxs-lookup"><span data-stu-id="bc031-114">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

* <span data-ttu-id="bc031-115">La dernière version de[Yeoman](https://yeoman.io/) et de [Yeoman Générateur de compléments Office](https://www.npmjs.com/package/generator-office). Pour installer ces outils globalement, exécutez la commande suivante à partir de l’invite de commande :</span><span class="sxs-lookup"><span data-stu-id="bc031-115">The latest version of [Yeoman](https://yeoman.io/) and the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > <span data-ttu-id="bc031-116">Même si vous avez précédemment installé la Yeoman générateur, nous vous recommandons une mise à jour de votre package à partir de la dernière version de npm.</span><span class="sxs-lookup"><span data-stu-id="bc031-116">Even if you have previously installed the Yeoman generator, we recommend updating your package to the latest version from npm.</span></span>

* <span data-ttu-id="bc031-117">Excel pour Windows (version 64 bits 1810 ou ultérieure) ou Excel Online</span><span class="sxs-lookup"><span data-stu-id="bc031-117">Excel for Windows (64-bit version 1810 or later) or Excel Online</span></span>

* <span data-ttu-id="bc031-118">Rejoignez le[programme Office Insider](https://products.office.com/office-insider)(\*\* niveau\*\*Insider, anciennement appelé « Insider Fast »)</span><span class="sxs-lookup"><span data-stu-id="bc031-118">Join the [Office Insider program](https://products.office.com/office-insider) (**Insider** level -- formerly called "Insider Fast")</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="bc031-119">Créer un projet de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="bc031-119">Create a custom functions project</span></span>

 <span data-ttu-id="bc031-120">Pour commencer, vous devez créer le projet de code pour créer votre complément de fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="bc031-120">To start, you'll create the code project to build your custom function add-in.</span></span> <span data-ttu-id="bc031-121">Le [ générateur Yeoman de compléments Office](https://www.npmjs.com/package/generator-office) permettront de configurer votre projet avec certaines fonctions personnalisées initiales que vous pouvez essayer.</span><span class="sxs-lookup"><span data-stu-id="bc031-121">The [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) will set up your project with some initial custom functions that you can try out.</span></span>

1. <span data-ttu-id="bc031-122">Exécutez la commande suivante, puis répondez aux invitations comme suit.</span><span class="sxs-lookup"><span data-stu-id="bc031-122">Run the following command and then answer the prompts as follows.</span></span>
    
    ```
    yo office
    ```
    
    * <span data-ttu-id="bc031-123">Choisissez un type de projet : `Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="bc031-123">Choose a project type: `Excel Custom Functions Add-in project (...)`</span></span>
    * <span data-ttu-id="bc031-124">Choisissez un type de script : `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="bc031-124">Choose a script type: `JavaScript`</span></span>
    * <span data-ttu-id="bc031-125">Comment souhaitez-vous nommer votre complément ?</span><span class="sxs-lookup"><span data-stu-id="bc031-125">What do you want to name your add-in?</span></span> `stock-ticker`
    
    ![Le générateur de yeoman pour les compléments Office vous invite pour les fonctions personnalisées](../images/12-10-fork-cf-pic.jpg)
    
    <span data-ttu-id="bc031-127">Le générateur Yeoman crée le projet et installe les composants Node.js de la prise en charge.</span><span class="sxs-lookup"><span data-stu-id="bc031-127">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="bc031-128">Accédez au dossier du projet.</span><span class="sxs-lookup"><span data-stu-id="bc031-128">Go to the project folder.</span></span>
    
    ```
    cd stock-ticker
    ```

3. <span data-ttu-id="bc031-129">Approuver le certificat auto-signé est nécessaire pour exécuter ce projet.</span><span class="sxs-lookup"><span data-stu-id="bc031-129">Trust the self-signed certificate that is needed to run this project.</span></span> <span data-ttu-id="bc031-130">Pour obtenir des instructions détaillées pour Windows ou Mac, voir [Ajout des Certificats Auto-signés comme Certificat Racine Approuvé](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span><span class="sxs-lookup"><span data-stu-id="bc031-130">For detailed instructions for either Windows or Mac, see [Adding Self Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span>  

4. <span data-ttu-id="bc031-131">Construire le projet.</span><span class="sxs-lookup"><span data-stu-id="bc031-131">Build the project.</span></span>
    
    ```
    npm run build
    ```

5. <span data-ttu-id="bc031-132">Démarrez le serveur web local qui est exécuté dans Node.js.</span><span class="sxs-lookup"><span data-stu-id="bc031-132">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="bc031-133">Vous pouvez tester le complément de fonction personnalisée dans Excel pour Windows ou Excel Online.</span><span class="sxs-lookup"><span data-stu-id="bc031-133">You can try out the custom function add-in in Excel for Windows, or Excel Online.</span></span>

# <a name="excel-for-windowstabexcel-windows"></a>[<span data-ttu-id="bc031-134">Excel pour Windows</span><span class="sxs-lookup"><span data-stu-id="bc031-134">Excel for Windows</span></span>](#tab/excel-windows)

<span data-ttu-id="bc031-135">Exécutez la commande suivante.</span><span class="sxs-lookup"><span data-stu-id="bc031-135">Run the following command:</span></span>

```
npm run start
```

<span data-ttu-id="bc031-136">Cette commande démarre le serveur web et le complément sideloads de votre fonction personnalisée dans Excel pour Windows.</span><span class="sxs-lookup"><span data-stu-id="bc031-136">This command starts the web server, and sideloads your custom function add-in into Excel for Windows.</span></span>

> [!NOTE]
> <span data-ttu-id="bc031-137">Si vous complément ne charge pas, vérifiez que vous avez correctement terminé l’étape 3.</span><span class="sxs-lookup"><span data-stu-id="bc031-137">If you add-in does not load, check that you have completed step 3 properly.</span></span>

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="bc031-138">Excel Online</span><span class="sxs-lookup"><span data-stu-id="bc031-138">Excel Online</span></span>](#tab/excel-online)

<span data-ttu-id="bc031-139">Exécutez la commande suivante.</span><span class="sxs-lookup"><span data-stu-id="bc031-139">Run the following command:</span></span>

```
npm run start-web
```

<span data-ttu-id="bc031-140">Cette commande démarre le service web.</span><span class="sxs-lookup"><span data-stu-id="bc031-140">This command starts the web server.</span></span> <span data-ttu-id="bc031-141">Procédez comme suit pour votre complément sideload.</span><span class="sxs-lookup"><span data-stu-id="bc031-141">Use the following steps to sideload your add-in.</span></span>

<ol type="a">
   <li><span data-ttu-id="bc031-142">Dans Excel Online, sélectionnez l’onglet <strong>Insérer</strong>, puis <strong>Compléments</strong>.</span><span class="sxs-lookup"><span data-stu-id="bc031-142">In Excel Online, choose the <strong>Insert</strong> tab and then choose <strong>Add-ins</strong>.  Insert ribbon in Excel Online with the My Add-ins icon highlighted</span></span><br/>
   <img src="../images/excel-cf-online-register-add-in-1.png" alt="Insert ribbon in Excel Online with the My Add-ins icon highlighted"></li>
   <li><span data-ttu-id="bc031-143">Sélectionnez<strong>Gérer mes Compléments</strong> et sélectionnez <strong>Télécharger mon complément</strong>.</span><span class="sxs-lookup"><span data-stu-id="bc031-143">Choose <strong>Manage My Add-ins</strong> and select <strong>Upload My Add-in</strong>.</span></span></li> 
   <li><span data-ttu-id="bc031-144">Sélectionnez <strong>Parcourir... </strong> et accédez au répertoire racine du projet créé par le Générateur de Yo Office.</span><span class="sxs-lookup"><span data-stu-id="bc031-144">Choose <strong>Browse...</strong> and navigate to the root directory of the project that the Yeoman generator created.</span></span></li> 
   <li><span data-ttu-id="bc031-145">Sélectionnez le fichier<strong>manifest.xml</strong> puis sélectionnez<strong>Ouvrir</strong>, puis sélectionnez <strong>Télécharger</strong>.</span><span class="sxs-lookup"><span data-stu-id="bc031-145">Select the file <strong>manifest.xml</strong> and choose <strong>Open</strong>, then choose <strong>Upload</strong>.</span></span></li>
</ol>

> [!NOTE]
> <span data-ttu-id="bc031-146">Si vous complément ne charge pas, vérifiez que vous avez correctement terminé l’étape 3.</span><span class="sxs-lookup"><span data-stu-id="bc031-146">If you add-in does not load, check that you have completed step 3 properly.</span></span>

--- 
    
## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="bc031-147">Essayer une fonction personnalisée prédéfinie</span><span class="sxs-lookup"><span data-stu-id="bc031-147">Try out a prebuilt custom function</span></span>

<span data-ttu-id="bc031-148">Le projet de fonctions personnalisées que vous avez créé déjà comporte deux fonctions personnalisées prédéfinies nommées AJOUTER et INCRÉMENT.</span><span class="sxs-lookup"><span data-stu-id="bc031-148">The custom functions project that you created alrady has two prebuilt custom functions named ADD and INCREMENT.</span></span> <span data-ttu-id="bc031-149">Le code de ces fonctions prédéfinis participe le fichier**src/customfunctions.js**.</span><span class="sxs-lookup"><span data-stu-id="bc031-149">The code for these prebuilt functions is in the  **src/customfunctions.js** file.</span></span> <span data-ttu-id="bc031-150">Le fichier**manifest.xml**indique que toutes les fonctions personnalisées appartiennent à l’`CONTOSO`espace de noms.</span><span class="sxs-lookup"><span data-stu-id="bc031-150">The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span> <span data-ttu-id="bc031-151">L’espace de noms CONTOSO permet d’accéder aux fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="bc031-151">You'll use the CONTOSO namespace to access the custom functions in Excel.</span></span>

<span data-ttu-id="bc031-152">Essayez de reproduire la`ADD` fonction personnalisée en complétant les étapes suivantes dans Excel:</span><span class="sxs-lookup"><span data-stu-id="bc031-152">In your Excel workbook, try out the `ADD` custom function by completing the following steps in Excel:</span></span>

1. <span data-ttu-id="bc031-153">Dans Excel, accédez à n’importe quelle cellule et entrez `=CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="bc031-153">In Excel, go to any cell and enter `=CONTOSO`.</span></span> <span data-ttu-id="bc031-154">Notez que le menu de saisie semi-automatique affiche la liste de toutes les fonctions dans l’espace de noms `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="bc031-154">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="bc031-155">Exécutez la`CONTOSO.ADD` fonction, avec les nombres `10` et `200` comme paramètres d’entrée, en spécifiant la valeur`=CONTOSO.ADD(10,200)`suivante dans la cellule et appuyez sur entrée.</span><span class="sxs-lookup"><span data-stu-id="bc031-155">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="bc031-156">Le `ADD` fonction personnalisée calcule la somme des deux nombres que vous avez spécifiés et renvoie le résultat**210** .</span><span class="sxs-lookup"><span data-stu-id="bc031-156">The `ADD` custom function computes the sum of the two numbers that you provided and returns the result of **210**.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="bc031-157">Créer une fonction personnalisée qui demande les données à partir du web</span><span class="sxs-lookup"><span data-stu-id="bc031-157">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="bc031-158">Intégration de données à partir du Web est un excellent moyen pour étendre Excel via les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="bc031-158">Integrating data from the Web is a great way to extend Excel through custom functions.</span></span> <span data-ttu-id="bc031-159">Vous allez ensuite créer une fonction personnalisée nommée `stockPrice` qui obtient des actions à partir d’une API Web et renvoie le résultat à la cellule d’une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="bc031-159">Next you’ll create a custom function named `stockPrice` that gets a stock quote from a Web API and returns the result to the cell of a worksheet.</span></span> <span data-ttu-id="bc031-160">Cette fonction personnalisée utilise l’API de cotation IEX, qui est gratuit et ne requiert pas d’authentification.</span><span class="sxs-lookup"><span data-stu-id="bc031-160">This custom function uses the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="bc031-161">Dans le projet**Bourse**, recherchez le fichier**src/customfunctions.js** et ouvrez-le dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="bc031-161">In the **stock-ticker** project that the Yeoman generator created, find the file **src/customfunctions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="bc031-162">Dans**customfunctions.js**, recherchez la`increment` fonction et ajoutez le code suivant immédiatement après cette fonction.</span><span class="sxs-lookup"><span data-stu-id="bc031-162">In **customfunctions.js**, locate the `increment` function and add the following code immediately after that function.</span></span>

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

3. In **customfunctions.js**, locate the line `CustomFunctions.associate("INCREMENT", increment);`. Add the following line of code immediately after that line, and save the file.

    ```js
    CustomFunctions.associate("STOCKPRICE", stockprice);
    ```

    <span data-ttu-id="bc031-163">Le `CustomFunctions.associate` code associe le `id`de la fonction avec l’adresse de la fonction de `increment` dans JavaScript afin qu’Excel peut appeler votre fonction.</span><span class="sxs-lookup"><span data-stu-id="bc031-163">The `CustomFunctions.associate` code associates the `id` of the function with the function address of `increment` in JavaScript so that Excel can call your function.</span></span>

    <span data-ttu-id="bc031-164">Avant qu’Excel puisse utiliser votre fonction personnalisée, vous devez le décrire utilisant des métadonnées.</span><span class="sxs-lookup"><span data-stu-id="bc031-164">Before Excel can use your custom function, you need to describe it using metadata.</span></span> <span data-ttu-id="bc031-165">Vous devez d’abord définir la méthode`id` utilisés dans le `associate`, ainsi que certaines autres métadonnées.</span><span class="sxs-lookup"><span data-stu-id="bc031-165">You need to define the `id` used in the `associate` method previously, along with some other metadata.</span></span>


4. <span data-ttu-id="bc031-166">Ouvrez le fichier**config/customfunctions.json**.</span><span class="sxs-lookup"><span data-stu-id="bc031-166">Open the **config/customfunctions.json** file.</span></span> <span data-ttu-id="bc031-167">Ajoutez l’objet JSON suivante à la matrice « fonctions » et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="bc031-167">Add the following JSON object to the 'functions' array and save the file.</span></span>

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
                "description": "stock symbol",
                "type": "string",
                "dimensionality": "scalar"
            }
        ]
    }
    ```

    <span data-ttu-id="bc031-168">Cet objet JSON décrit la fonction`stockPrice`, ses paramètres, et le type de résultat qu’il renvoie.</span><span class="sxs-lookup"><span data-stu-id="bc031-168">This JSON describes the `stockPrice` function, its parameters, and the type of result it returns.</span></span>

5. <span data-ttu-id="bc031-169">Enregistrez de nouveau le complément dans Excel afin que la nouvelle fonction soit disponible.</span><span class="sxs-lookup"><span data-stu-id="bc031-169">Re-register the add-in in Excel so that the new function is available.</span></span> 

# <a name="excel-for-windowstabexcel-windows"></a>[<span data-ttu-id="bc031-170">Excel pour Windows</span><span class="sxs-lookup"><span data-stu-id="bc031-170">Excel for Windows</span></span>](#tab/excel-windows)

1. <span data-ttu-id="bc031-171">Fermez Excel, puis ouvrez de nouveau Excel.</span><span class="sxs-lookup"><span data-stu-id="bc031-171">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="bc031-172">Dans Excel, sélectionnez l’onglet**insérer**, puis cliquez sur la flèche vers le bas située à droite de **mes compléments**.  ![Insérer du ruban dans Excel pour Windows avec la flèche Mes compléments mise en évidence](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="bc031-172">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

3. <span data-ttu-id="bc031-173">Dans la liste des compléments disponibles, recherchez la section **Compléments Développeur** et sélectionnez votre complément**bourse** pour effectuer cette opération.</span><span class="sxs-lookup"><span data-stu-id="bc031-173">In the list of available add-ins, find the **Developer Add-ins** section and select the **stock-ticker** add-in to register it.</span></span>
    <span data-ttu-id="bc031-174">![Insérer du ruban dans Excel pour Windows avec le complément Fonctions personnalisées Excel mis en évidence dans la liste Mes compléments](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="bc031-174">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="bc031-175">Excel Online</span><span class="sxs-lookup"><span data-stu-id="bc031-175">Excel Online</span></span>](#tab/excel-online)

1. <span data-ttu-id="bc031-176">Dans Excel Online, sélectionnez l’onglet **insérer**, puis **compléments**. ![Insérer du ruban dans Excel Online avec l’icône Mes compléments mis en évidence](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="bc031-176">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="bc031-177">Sélectionnez**Gérer mes compléments** et sélectionnez **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="bc031-177">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

3. <span data-ttu-id="bc031-178">Sélectionnez \*\*Parcourir... \*\* et accédez au répertoire racine du projet créé par le Générateur de Yo Office.</span><span class="sxs-lookup"><span data-stu-id="bc031-178">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span> 

4. <span data-ttu-id="bc031-179">Sélectionnez le fichier**manifest.xml** puis sélectionnez**Ouvrir**, puis sélectionnez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="bc031-179">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 

<ol start="6">
<li> <span data-ttu-id="bc031-180">Essayez la nouvelle fonction.</span><span class="sxs-lookup"><span data-stu-id="bc031-180">Try out the new function.</span></span> <span data-ttu-id="bc031-181">Dans la cellule <strong>B1</strong>, tapez le texte <strong>= CONTOSO. STOCKPRICE("MSFT")</strong> et appuyez sur ENTRÉE.</span><span class="sxs-lookup"><span data-stu-id="bc031-181">In cell <strong>B1</strong>, type the text <strong></strong> and press enter.</span></span> <span data-ttu-id="bc031-182">Vous devriez voir que le résultat dans la cellule <strong>B1</strong> est le prix boursier actuel pour un partage de stock Microsoft.</span><span class="sxs-lookup"><span data-stu-id="bc031-182">You should see that the result in cell <strong>B1</strong> is the current stock price for one share of Microsoft stock.</span></span></li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="bc031-183">Créer une fonction personnalisée asynchrone diffusion en continu</span><span class="sxs-lookup"><span data-stu-id="bc031-183">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="bc031-184">La fonction`stockPrice`que vous venez de créer renvoie le prix d’une action à un moment donné, mais les prix des actions changent constamment.</span><span class="sxs-lookup"><span data-stu-id="bc031-184">The `stockPrice` function that you just created returns the price of a stock at a specific moment in time, but stock prices are always changing.</span></span> <span data-ttu-id="bc031-185">Vous allez ensuite créer une fonction personnalisée nommée `stockPriceStream` qui obtient le prix d’une action chaque 1000 millisecondes.</span><span class="sxs-lookup"><span data-stu-id="bc031-185">Next you’ll create a custom function named `stockPriceStream` that gets the price of a stock every 1000 milliseconds.</span></span>

1. <span data-ttu-id="bc031-186">Dans le projet **Bourse**, ajoutez le fichier **src/customfunctions.js** et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="bc031-186">In the **stock-ticker** project that the Yeoman generator created, add the following code to **src/customfunctions.js** and save the file.</span></span>

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
    
    CustomFunctions.associate("STOCKPRICESTREAM", stockpricestream);
    ```
    
    <span data-ttu-id="bc031-187">Avant qu’Excel puisse utiliser votre fonction personnalisée, vous devez le décrire utilisant des métadonnées.</span><span class="sxs-lookup"><span data-stu-id="bc031-187">Before Excel can use your custom function, you need to describe it using metadata.</span></span>
    
2. <span data-ttu-id="bc031-188">Dans le projet**bourse** ajoutez l’objet suivant à la `functions` matrice au sein du fichier **config/customfunctions.json** et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="bc031-188">In the **stock-ticker** project that the Yeoman generator created, add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>
    
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
                "description": "stock symbol",
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

    <span data-ttu-id="bc031-189">Cet élément JSON décrit la fonction`stockPriceStream`.</span><span class="sxs-lookup"><span data-stu-id="bc031-189">This JSON describes the `stockPriceStream` function.</span></span> <span data-ttu-id="bc031-190">Pour n’importe quelle fonction de diffusion en continu, la propriété`stream` et la propriété`cancelable`doivent être définies `true` au sein de l’ `options` objet, comme illustré dans cet exemple de code.</span><span class="sxs-lookup"><span data-stu-id="bc031-190">For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.</span></span>

3. <span data-ttu-id="bc031-191">Enregistrez de nouveau le complément dans Excel afin que la nouvelle fonction soit disponible.</span><span class="sxs-lookup"><span data-stu-id="bc031-191">Re-register the add-in in Excel so that the new function is available.</span></span>

# <a name="excel-for-windowstabexcel-windows"></a>[<span data-ttu-id="bc031-192">Excel pour Windows</span><span class="sxs-lookup"><span data-stu-id="bc031-192">Excel for Windows</span></span>](#tab/excel-windows)

1. <span data-ttu-id="bc031-193">Fermez Excel, puis ouvrez de nouveau Excel.</span><span class="sxs-lookup"><span data-stu-id="bc031-193">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="bc031-194">Dans Excel, sélectionnez l’onglet**insérer**, puis cliquez sur la flèche vers le bas située à droite de **mes compléments**.  ![Insérer du ruban dans Excel pour Windows avec la flèche Mes compléments mise en évidence](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="bc031-194">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

3. <span data-ttu-id="bc031-195">Dans la liste des compléments disponibles, recherchez la section **Compléments Développeur** et sélectionnez votre complément**bourse** pour effectuer cette opération.</span><span class="sxs-lookup"><span data-stu-id="bc031-195">In the list of available add-ins, find the **Developer Add-ins** section and select the **stock-ticker** add-in to register it.</span></span>
    <span data-ttu-id="bc031-196">![Insérer du ruban dans Excel pour Windows avec le complément Fonctions personnalisées Excel mis en évidence dans la liste Mes compléments](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="bc031-196">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="bc031-197">Excel Online</span><span class="sxs-lookup"><span data-stu-id="bc031-197">Excel Online</span></span>](#tab/excel-online)

1. <span data-ttu-id="bc031-198">Dans Excel Online, sélectionnez l’onglet **insérer**, puis **compléments**. ![Insérer du ruban dans Excel Online avec l’icône Mes compléments mis en évidence](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="bc031-198">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="bc031-199">Sélectionnez**Gérer mes compléments** et sélectionnez **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="bc031-199">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="bc031-200">Sélectionnez \*\*Parcourir... \*\* et accédez au répertoire racine du projet créé par le Générateur de Yo Office.</span><span class="sxs-lookup"><span data-stu-id="bc031-200">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="bc031-201">Sélectionnez le fichier**manifest.xml** puis sélectionnez**Ouvrir**, puis sélectionnez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="bc031-201">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 

<ol start="4">
<li><span data-ttu-id="bc031-202">Essayez la nouvelle fonction.</span><span class="sxs-lookup"><span data-stu-id="bc031-202">Try out the new function.</span></span> <span data-ttu-id="bc031-203">Dans la cellule <strong>C1</strong>, tapez le texte <strong>= CONTOSO. STOCKPRICE("MSFT")</strong> et appuyez sur ENTRÉE.</span><span class="sxs-lookup"><span data-stu-id="bc031-203">In cell <strong>C1</strong>, type the text <strong></strong> and press enter.</span></span> <span data-ttu-id="bc031-204">Si le marché est ouvert, vous devriez voir que le résultat dans la cellule <strong>C1</strong> constamment mis à jour pour refléter le prix en temps réel pour un partage d’actions Microsoft.</span><span class="sxs-lookup"><span data-stu-id="bc031-204">Provided that the stock market is open, you should see that the result in cell <strong>C1</strong> is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span></li>
</ol>

## <a name="next-steps"></a><span data-ttu-id="bc031-205">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="bc031-205">Next steps</span></span>

<span data-ttu-id="bc031-206">Félicitations !</span><span class="sxs-lookup"><span data-stu-id="bc031-206">Congratulations!</span></span> <span data-ttu-id="bc031-207">Vous avez créé un nouveau projet de fonctions personnalisées, essayé une fonction prédéfinie, créé une fonction personnalisée qui demande les données à partir du web et créé une fonction personnalisée qui diffuse les données en temps réel à partir du web.</span><span class="sxs-lookup"><span data-stu-id="bc031-207">In this tutorial, you've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web.</span></span> <span data-ttu-id="bc031-208">Pour en savoir plus sur les fonctions personnalisées dans Excel, passez à l’article suivant :</span><span class="sxs-lookup"><span data-stu-id="bc031-208">To learn more about custom functions in Excel, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="bc031-209">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="bc031-209">Create custom functions in Excel</span></span>](../excel/custom-functions-overview.md)

### <a name="legal-information"></a><span data-ttu-id="bc031-210">Informations légales</span><span class="sxs-lookup"><span data-stu-id="bc031-210">Legal information</span></span>

<span data-ttu-id="bc031-211">Données fournies gratuitement par [IEX](https://iextrading.com/developer/).</span><span class="sxs-lookup"><span data-stu-id="bc031-211">Data provided free by [IEX](https://iextrading.com/developer/).</span></span> <span data-ttu-id="bc031-212">Afficher les [conditions d’utilisation de IEX](https://iextrading.com/api-exhibit-a/).</span><span class="sxs-lookup"><span data-stu-id="bc031-212">View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/).</span></span> <span data-ttu-id="bc031-213">L’utilisation de Microsoft de l’API IEX dans ce didacticiel est uniquement à des fins d’enseignement.</span><span class="sxs-lookup"><span data-stu-id="bc031-213">Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>


