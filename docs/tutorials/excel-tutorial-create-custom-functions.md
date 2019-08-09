---
title: Didacticiel de fonctions personnalisées Excel
description: Dans ce didacticiel, vous allez créer un complément Excel qui contient une fonction personnalisée qui effectue des calculs, requiert des données web ou lance un flux de données web.
ms.date: 07/09/2019
ms.prod: excel
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: 4b74463eafd5ac1b70e59cef6ef1f9f33cf0ffa2
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268179"
---
# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="afa75-103">Didacticiel : créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="afa75-103">Tutorial: Create custom functions in Excel</span></span>

<span data-ttu-id="afa75-104">Les fonctions personnalisées vous permettent d’ajouter de nouvelles fonctions dans Excel en définissant ces fonctions dans JavaScript dans le cadre d’un complément.</span><span class="sxs-lookup"><span data-stu-id="afa75-104">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="afa75-105">Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe quelle fonction native d’Excel, telle que `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="afa75-105">Users within Excel can access custom functions as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="afa75-106">Vous pouvez créer des fonctions personnalisées qui effectuent des tâches simples comme des calculs personnalisés ou des tâches plus complexes telles que la diffusion en continu des données en temps réel à partir du web dans une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="afa75-106">You can create custom functions that perform simple tasks like calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="afa75-107">Dans ce didacticiel, vous allez :</span><span class="sxs-lookup"><span data-stu-id="afa75-107">In this tutorial, you will:</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="afa75-108">Créer un complément de fonction personnalisée à l’aide la [Générateur Yeoman de compléments Office](https://www.npmjs.com/package/generator-office).</span><span class="sxs-lookup"><span data-stu-id="afa75-108">Create a custom function add-in using the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office).</span></span> 
> * <span data-ttu-id="afa75-109">Utiliser une fonction personnalisée prédéfinie pour effectuer un calcul simple.</span><span class="sxs-lookup"><span data-stu-id="afa75-109">Use a prebuilt custom function to perform a simple calculation.</span></span>
> * <span data-ttu-id="afa75-110">Créer une fonction personnalisée qui demande les données à partir du web.</span><span class="sxs-lookup"><span data-stu-id="afa75-110">Create a custom function that gets data from the web.</span></span>
> * <span data-ttu-id="afa75-111">Créer une fonction personnalisée qui diffuse les données en temps réel à partir du web.</span><span class="sxs-lookup"><span data-stu-id="afa75-111">Create a custom function that streams real-time data from the web.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="afa75-112">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="afa75-112">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* <span data-ttu-id="afa75-113">Excel sur Windows (version 1904 ou ultérieure, connexion à l’abonnement Office 365) ou sur le Web</span><span class="sxs-lookup"><span data-stu-id="afa75-113">Excel on Windows (version 1904 or later, connected to Office 365 subscription) or on the web</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="afa75-114">Créer un projet de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="afa75-114">Create a custom functions project</span></span>

 <span data-ttu-id="afa75-115">Pour commencer, vous devez créer le projet de code pour créer votre complément de fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="afa75-115">To start, you'll create the code project to build your custom function add-in.</span></span> <span data-ttu-id="afa75-116">Le [Générateur Yeoman pour les compléments Office](https://www.npmjs.com/package/generator-office) configurera votre projet avec certaines fonctions personnalisées prédéfinies que vous pouvez tester. Si vous avez déjà exécuté le démarrage rapide des fonctions personnalisées et généré un projet, continuez à utiliser ce projet et passez à [cette étape](#create-a-custom-function-that-requests-data-from-the-web) .</span><span class="sxs-lookup"><span data-stu-id="afa75-116">The [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) will set up your project with some prebuilt custom functions that you can try out. If you have already run the custom functions quick start and generated a project, continue to use that project and skip to [this step](#create-a-custom-function-that-requests-data-from-the-web) instead.</span></span>

1. <span data-ttu-id="afa75-117">Exécutez la commande suivante, puis répondez aux invitations comme suit.</span><span class="sxs-lookup"><span data-stu-id="afa75-117">Run the following command and then answer the prompts as follows.</span></span>
    
    ```command&nbsp;line
    yo office
    ```
    
    * <span data-ttu-id="afa75-118">**Sélectionnez un type de projet :** `Excel Custom Functions Add-in project`</span><span class="sxs-lookup"><span data-stu-id="afa75-118">**Choose a project type:** `Excel Custom Functions Add-in project`</span></span>
    * <span data-ttu-id="afa75-119">**Sélectionnez un type de script :** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="afa75-119">**Choose a script type:** `JavaScript`</span></span>
    * <span data-ttu-id="afa75-120">**Comment souhaitez-vous nommer votre complément ?**</span><span class="sxs-lookup"><span data-stu-id="afa75-120">**What do you want to name your add-in?**</span></span> `starcount`

    ![Le générateur de yeoman pour les compléments Office vous invite pour les fonctions personnalisées](../images/starcountPrompt.png)
    
    <span data-ttu-id="afa75-122">Le générateur crée le projet et installe les composants Node.js de la prise en charge.</span><span class="sxs-lookup"><span data-stu-id="afa75-122">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="afa75-123">Accédez au dossier racine du projet.</span><span class="sxs-lookup"><span data-stu-id="afa75-123">Navigate to the root folder of the project.</span></span>
    
    ```command&nbsp;line
    cd starcount
    ```

3. <span data-ttu-id="afa75-124">Créez le projet.</span><span class="sxs-lookup"><span data-stu-id="afa75-124">Build the project.</span></span>
    
    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > <span data-ttu-id="afa75-125">Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez.</span><span class="sxs-lookup"><span data-stu-id="afa75-125">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="afa75-126">Si vous êtes invité à installer un certificat après avoir exécuté `npm run build`, acceptez d’installer le certificat fourni par le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="afa75-126">If you are prompted to install a certificate after you run `npm run build`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

4. <span data-ttu-id="afa75-127">Démarrez le serveur web local qui est exécuté dans Node.js.</span><span class="sxs-lookup"><span data-stu-id="afa75-127">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="afa75-128">Vous pouvez essayer le complément de fonction personnalisée dans Excel sur le Web ou Windows.</span><span class="sxs-lookup"><span data-stu-id="afa75-128">You can try out the custom function add-in in Excel on the web or Windows.</span></span>

# <a name="excel-on-windows-or-mactabexcel-windows"></a>[<span data-ttu-id="afa75-129">Excel sur Windows ou Mac</span><span class="sxs-lookup"><span data-stu-id="afa75-129">Excel on Windows or Mac</span></span>](#tab/excel-windows)

<span data-ttu-id="afa75-130">Pour tester votre complément dans Excel sous Windows ou Mac, exécutez la commande suivante:</span><span class="sxs-lookup"><span data-stu-id="afa75-130">To test your add-in in Excel on Windows or Mac, run the following command.</span></span> <span data-ttu-id="afa75-131">Lorsque vous exécutez cette commande, le serveur Web local démarre et Excel s’ouvre avec votre complément chargé.</span><span class="sxs-lookup"><span data-stu-id="afa75-131">When you run this command, the local web server will start and Excel will open with your add-in loaded.</span></span>

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-on-the-webtabexcel-online"></a>[<span data-ttu-id="afa75-132">Excel sur le Web</span><span class="sxs-lookup"><span data-stu-id="afa75-132">Excel on the web</span></span>](#tab/excel-online)

<span data-ttu-id="afa75-133">Pour tester votre complément dans Excel sur un navigateur, exécutez la commande suivante.</span><span class="sxs-lookup"><span data-stu-id="afa75-133">To test your add-in in Excel on a browser, run the following command.</span></span> <span data-ttu-id="afa75-134">Lorsque vous exécutez cette commande, le serveur web local démarre.</span><span class="sxs-lookup"><span data-stu-id="afa75-134">When you run this command, the local web server will start.</span></span>

```command&nbsp;line
npm run start:web
```

<span data-ttu-id="afa75-135">Pour utiliser votre complément de fonctions personnalisées, ouvrez un nouveau classeur dans Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="afa75-135">To use your custom functions add-in, open a new workbook in Excel on the web.</span></span> <span data-ttu-id="afa75-136">Dans ce classeur, effectuez les étapes suivantes pour chargement votre complément.</span><span class="sxs-lookup"><span data-stu-id="afa75-136">In this workbook, complete the following steps to sideload your add-in.</span></span>

1. <span data-ttu-id="afa75-137">Dans Excel, sélectionnez l’onglet **insertion** , puis **compléments**.</span><span class="sxs-lookup"><span data-stu-id="afa75-137">In Excel, choose the **Insert** tab and then choose **Add-ins**.</span></span>

   ![Insérer un ruban dans Excel sur le Web avec l’icône mes compléments mise en surbrillance](../images/excel-cf-online-register-add-in-1.png)
   
2. <span data-ttu-id="afa75-139">Sélectionnez**Gérer mes compléments** et sélectionnez **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="afa75-139">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="afa75-140">Sélectionnez \*\*Parcourir... \*\* et accédez au répertoire racine du projet créé par le Générateur de Yo Office.</span><span class="sxs-lookup"><span data-stu-id="afa75-140">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="afa75-141">Sélectionnez le fichier**manifest.xml** puis sélectionnez**Ouvrir**, puis sélectionnez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="afa75-141">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 
    
## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="afa75-142">Essayer une fonction personnalisée prédéfinie</span><span class="sxs-lookup"><span data-stu-id="afa75-142">Try out a prebuilt custom function</span></span>

<span data-ttu-id="afa75-143">Le projet de fonctions personnalisées que vous avez créé contient des fonctions personnalisées prédéfinies, définies dans le fichier **./SRC/Functions/functions.js** .</span><span class="sxs-lookup"><span data-stu-id="afa75-143">The custom functions project that you created contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file.</span></span> <span data-ttu-id="afa75-144">Le fichier**manifest.xml**indique que toutes les fonctions personnalisées appartiennent à l’`CONTOSO`espace de noms.</span><span class="sxs-lookup"><span data-stu-id="afa75-144">The **./manifest.xml** file specifies that all custom functions belong to the `CONTOSO` namespace.</span></span> <span data-ttu-id="afa75-145">L’espace de noms CONTOSO permet d’accéder aux fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="afa75-145">You'll use the CONTOSO namespace to access the custom functions in Excel.</span></span>

<span data-ttu-id="afa75-146">Essayez de reproduire la`ADD` fonction personnalisée en complétant les étapes suivantes dans Excel:</span><span class="sxs-lookup"><span data-stu-id="afa75-146">Next you'll try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="afa75-147">Dans Excel, accédez à n’importe quelle cellule et entrez `=CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="afa75-147">In Excel, go to any cell and enter `=CONTOSO`.</span></span> <span data-ttu-id="afa75-148">Notez que le menu de saisie semi-automatique affiche la liste de toutes les fonctions dans l’espace de noms `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="afa75-148">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="afa75-149">Exécutez la`CONTOSO.ADD` fonction, avec les nombres `10` et `200` comme paramètres d’entrée, en spécifiant la valeur`=CONTOSO.ADD(10,200)`suivante dans la cellule et appuyez sur entrée.</span><span class="sxs-lookup"><span data-stu-id="afa75-149">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="afa75-150">Le `ADD` fonction personnalisée calcule la somme des deux nombres que vous avez spécifiés et renvoie le résultat**210** .</span><span class="sxs-lookup"><span data-stu-id="afa75-150">The `ADD` custom function computes the sum of the two numbers that you provided and returns the result of **210**.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="afa75-151">Créer une fonction personnalisée qui demande les données à partir du web</span><span class="sxs-lookup"><span data-stu-id="afa75-151">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="afa75-152">Intégration de données à partir du Web est un excellent moyen pour étendre Excel via les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="afa75-152">Integrating data from the Web is a great way to extend Excel through custom functions.</span></span> <span data-ttu-id="afa75-153">Ensuite, vous allez créer une fonction personnalisée `getStarCount` nommée qui indique le nombre d’étoiles dont dispose un référentiel GitHub donné.</span><span class="sxs-lookup"><span data-stu-id="afa75-153">Next you’ll create a custom function named `getStarCount` that shows how many stars a given Github repository possesses.</span></span>

1. <span data-ttu-id="afa75-154">Dans le projet **starcount** , recherchez le fichier **./SRC/Functions/functions.js** et ouvrez-le dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="afa75-154">In the **starcount** project, find the file **./src/functions/functions.js** and open it in your code editor.</span></span> 

2. <span data-ttu-id="afa75-155">Dans **Function. js**, ajoutez le code suivant:</span><span class="sxs-lookup"><span data-stu-id="afa75-155">In **function.js**, add the following code:</span></span> 

```JS
/**
  * Gets the star count for a given Github repository.
  * @customfunction 
  * @param {string} userName string name of Github user or organization.
  * @param {string} repoName string name of the Github repository.
  * @return {number} number of stars given to a Github repository.
  */
  async function getStarCount(userName, repoName) {
    try {
      //You can change this URL to any web request you want to work with.
      const url = "https://api.github.com/repos/" + userName + "/" + repoName;
      const response = await fetch(url);
      //Expect that status code is in 200-299 range
      if (!response.ok) {
        throw new Error(response.statusText)
      }
        const jsonResponse = await response.json();
        return jsonResponse.watchers_count;
    }
    catch (error) {
      return error;
    }
  }
```

3. <span data-ttu-id="afa75-156">Exécutez la commande suivante pour regénérer le projet.</span><span class="sxs-lookup"><span data-stu-id="afa75-156">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

4. <span data-ttu-id="afa75-157">Procédez comme suit (pour Excel sur le Web, Windows ou Mac) pour réenregistrer le complément dans Excel.</span><span class="sxs-lookup"><span data-stu-id="afa75-157">Complete the following steps (for Excel on the web, Windows, or Mac) to re-register the add-in in Excel.</span></span> <span data-ttu-id="afa75-158">Vous devez effectuer ces étapes avant que la nouvelle fonction ne soit disponible.</span><span class="sxs-lookup"><span data-stu-id="afa75-158">You must complete these steps before the new function will be available.</span></span>

### <a name="excel-on-windows-or-mactabexcel-windows"></a>[<span data-ttu-id="afa75-159">Excel sur Windows ou Mac</span><span class="sxs-lookup"><span data-stu-id="afa75-159">Excel on Windows or Mac</span></span>](#tab/excel-windows)

1. <span data-ttu-id="afa75-160">Fermez Excel, puis ouvrez de nouveau Excel.</span><span class="sxs-lookup"><span data-stu-id="afa75-160">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="afa75-161">Dans Excel, sélectionnez l’onglet **Insérer** , puis cliquez sur la flèche vers le bas située à droite de **mes compléments**.  ![Insérer un ruban dans Excel sur Windows avec la flèche mes compléments mise en surbrillance](../images/select-insert.png)</span><span class="sxs-lookup"><span data-stu-id="afa75-161">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel on Windows with the My Add-ins arrow highlighted](../images/select-insert.png)</span></span>

3. <span data-ttu-id="afa75-162">Dans la liste des compléments disponibles, recherchez la section **compléments pour développeurs** et sélectionnez le complément **starcount** pour l’enregistrer.</span><span class="sxs-lookup"><span data-stu-id="afa75-162">In the list of available add-ins, find the **Developer Add-ins** section and select the **starcount** add-in to register it.</span></span>
    <span data-ttu-id="afa75-163">![Insérer un ruban dans Excel sur Windows avec le complément de fonctions personnalisées Excel mis en surbrillance dans la liste mes compléments](../images/list-starcount.png)</span><span class="sxs-lookup"><span data-stu-id="afa75-163">![Insert ribbon in Excel on Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/list-starcount.png)</span></span>


# <a name="excel-on-the-webtabexcel-online"></a>[<span data-ttu-id="afa75-164">Excel sur le Web</span><span class="sxs-lookup"><span data-stu-id="afa75-164">Excel on the web</span></span>](#tab/excel-online)

1. <span data-ttu-id="afa75-165">Dans Excel, sélectionnez l’onglet **insertion** , puis **compléments**.  ![Insérer un ruban dans Excel sur le Web avec l’icône mes compléments mise en surbrillance](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="afa75-165">In Excel, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel on the web with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="afa75-166">Sélectionnez**Gérer mes compléments** et sélectionnez **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="afa75-166">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="afa75-167">Sélectionnez \*\*Parcourir... \*\* et accédez au répertoire racine du projet créé par le Générateur de Yo Office.</span><span class="sxs-lookup"><span data-stu-id="afa75-167">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="afa75-168">Sélectionnez le fichier**manifest.xml** puis sélectionnez**Ouvrir**, puis sélectionnez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="afa75-168">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

<ol start="5">
<li> <span data-ttu-id="afa75-169">Essayez la nouvelle fonction.</span><span class="sxs-lookup"><span data-stu-id="afa75-169">Try out the new function.</span></span> <span data-ttu-id="afa75-170">Dans la cellule <strong>B1</strong>, tapez le texte <strong>= contoso. GETSTARCOUNT ("OfficeDev", "Excel-Custom-Functions")</strong> et appuyez sur entrée.</span><span class="sxs-lookup"><span data-stu-id="afa75-170">In cell <strong>B1</strong>, type the text <strong>=CONTOSO.GETSTARCOUNT("OfficeDev", "Excel-Custom-Functions")</strong> and press enter.</span></span> <span data-ttu-id="afa75-171">Vous devriez voir que le résultat dans la cellule <strong>B1</strong> est le nombre actuel d’étoiles attribuées au [référentiel GitHub de fonctions personnalisées Excel](https://github.com/OfficeDev/Excel-Custom-Functions).</span><span class="sxs-lookup"><span data-stu-id="afa75-171">You should see that the result in cell <strong>B1</strong> is the current number of stars given to the [Excel-Custom-Functions Github repository](https://github.com/OfficeDev/Excel-Custom-Functions).</span></span></li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="afa75-172">Créer une fonction personnalisée asynchrone diffusion en continu</span><span class="sxs-lookup"><span data-stu-id="afa75-172">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="afa75-173">La `getStarCount` fonction renvoie le nombre d’étoiles qu’un référentiel a à un moment donné.</span><span class="sxs-lookup"><span data-stu-id="afa75-173">The `getStarCount` function returns the number of stars a repository has at a specific moment in time.</span></span> <span data-ttu-id="afa75-174">Les fonctions personnalisées peuvent également renvoyer des données qui changent en permanence.</span><span class="sxs-lookup"><span data-stu-id="afa75-174">Custom functions can also return data that is continuously changing.</span></span> <span data-ttu-id="afa75-175">Ces fonctions sont appelées fonctions de diffusion en continu.</span><span class="sxs-lookup"><span data-stu-id="afa75-175">These functions are called streaming functions.</span></span> <span data-ttu-id="afa75-176">Elles doivent inclure un `invocation` paramètre qui fait référence à la cellule à partir de laquelle la fonction a été appelée.</span><span class="sxs-lookup"><span data-stu-id="afa75-176">They must include an `invocation` parameter which refers to the cell where the function was called from.</span></span> <span data-ttu-id="afa75-177">Le `invocation` paramètre est utilisé pour mettre à jour le contenu de la cellule à tout moment.</span><span class="sxs-lookup"><span data-stu-id="afa75-177">The `invocation` parameter is used to update the contents of the cell at any time.</span></span>  

<span data-ttu-id="afa75-178">Dans l’exemple de code suivant, vous remarquerez qu’il existe deux `currentTime` fonctions `clock`, et.</span><span class="sxs-lookup"><span data-stu-id="afa75-178">In the following code sample, you'll notice that there are two functions, `currentTime` and `clock`.</span></span> <span data-ttu-id="afa75-179">La `currentTime` fonction est une fonction statique qui n’utilise pas la diffusion en continu.</span><span class="sxs-lookup"><span data-stu-id="afa75-179">The `currentTime` function is a static function that does not use streaming.</span></span> <span data-ttu-id="afa75-180">Elle renvoie la date sous la forme d’une chaîne.</span><span class="sxs-lookup"><span data-stu-id="afa75-180">It returns the date as a string.</span></span> <span data-ttu-id="afa75-181">La `clock` fonction utilise la `currentTime` fonction pour fournir la nouvelle fois toutes les secondes à une cellule dans Excel.</span><span class="sxs-lookup"><span data-stu-id="afa75-181">The `clock` function uses the `currentTime` function to provide the new time every second to a cell in Excel.</span></span> <span data-ttu-id="afa75-182">Il utilise `invocation.setResult` pour fournir le temps à la cellule Excel et `invocation.onCanceled` pour gérer ce qui se produit lorsque la fonction est annulée.</span><span class="sxs-lookup"><span data-stu-id="afa75-182">It uses `invocation.setResult` to deliver the time to the Excel cell and `invocation.onCanceled` to handle what occurs when the function is canceled.</span></span>

1. <span data-ttu-id="afa75-183">Dans le projet **starcount** , ajoutez le code suivant à **./SRC/Functions/functions.js** et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="afa75-183">In the **starcount** project, add the following code to **./src/functions/functions.js** and save the file.</span></span>

```JS
/**
 * Returns the current time
 * @returns {string} String with the current time formatted for the current locale.
 */
function currentTime() {
  return new Date().toLocaleTimeString();
}

 /**
 * Displays the current time once a second
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 */
function clock(invocation) {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}
```

2. <span data-ttu-id="afa75-184">Exécutez la commande suivante pour regénérer le projet.</span><span class="sxs-lookup"><span data-stu-id="afa75-184">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

3. <span data-ttu-id="afa75-185">Procédez comme suit (pour Excel sur le Web, Windows ou Mac) pour réenregistrer le complément dans Excel.</span><span class="sxs-lookup"><span data-stu-id="afa75-185">Complete the following steps (for Excel on the web, Windows, or Mac) to re-register the add-in in Excel.</span></span> <span data-ttu-id="afa75-186">Vous devez effectuer ces étapes avant que la nouvelle fonction ne soit disponible.</span><span class="sxs-lookup"><span data-stu-id="afa75-186">You must complete these steps before the new function will be available.</span></span> 

# <a name="excel-on-windows-or-mactabexcel-windows"></a>[<span data-ttu-id="afa75-187">Excel sur Windows ou Mac</span><span class="sxs-lookup"><span data-stu-id="afa75-187">Excel on Windows or Mac</span></span>](#tab/excel-windows)

1. <span data-ttu-id="afa75-188">Fermez Excel, puis ouvrez de nouveau Excel.</span><span class="sxs-lookup"><span data-stu-id="afa75-188">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="afa75-189">Dans Excel, sélectionnez l’onglet **Insérer** , puis cliquez sur la flèche vers le bas située à droite de **mes compléments**.  ![Insérer un ruban dans Excel sur Windows avec la flèche mes compléments mise en surbrillance](../images/select-insert.png)</span><span class="sxs-lookup"><span data-stu-id="afa75-189">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel on Windows with the My Add-ins arrow highlighted](../images/select-insert.png)</span></span>

3. <span data-ttu-id="afa75-190">Dans la liste des compléments disponibles, recherchez la section **compléments pour développeurs** et sélectionnez le complément **starcount** pour l’enregistrer.</span><span class="sxs-lookup"><span data-stu-id="afa75-190">In the list of available add-ins, find the **Developer Add-ins** section and select the **starcount** add-in to register it.</span></span>
    <span data-ttu-id="afa75-191">![Insérer un ruban dans Excel sur Windows avec le complément de fonctions personnalisées Excel mis en surbrillance dans la liste mes compléments](../images/list-starcount.png)</span><span class="sxs-lookup"><span data-stu-id="afa75-191">![Insert ribbon in Excel on Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/list-starcount.png)</span></span>

# <a name="excel-on-the-webtabexcel-online"></a>[<span data-ttu-id="afa75-192">Excel sur le Web</span><span class="sxs-lookup"><span data-stu-id="afa75-192">Excel on the web</span></span>](#tab/excel-online)

1. <span data-ttu-id="afa75-193">Dans Excel, sélectionnez l’onglet **insertion** , puis **compléments**.  ![Insérer un ruban dans Excel sur le Web avec l’icône mes compléments mise en surbrillance](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="afa75-193">In Excel, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel on the web with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="afa75-194">Sélectionnez**Gérer mes compléments** et sélectionnez **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="afa75-194">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="afa75-195">Sélectionnez \*\*Parcourir... \*\* et accédez au répertoire racine du projet créé par le Générateur de Yo Office.</span><span class="sxs-lookup"><span data-stu-id="afa75-195">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="afa75-196">Sélectionnez le fichier**manifest.xml** puis sélectionnez**Ouvrir**, puis sélectionnez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="afa75-196">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 

<ol start="4">
<li><span data-ttu-id="afa75-197">Essayez la nouvelle fonction.</span><span class="sxs-lookup"><span data-stu-id="afa75-197">Try out the new function.</span></span> <span data-ttu-id="afa75-198">Dans la cellule <strong>C1</strong>, tapez le texte <strong>= contoso. CLOCK ())</strong> , puis appuyez sur entrée.</span><span class="sxs-lookup"><span data-stu-id="afa75-198">In cell <strong>C1</strong>, type the text <strong>=CONTOSO.CLOCK())</strong> and press enter.</span></span> <span data-ttu-id="afa75-199">Vous devriez voir la date du jour, qui diffuse une mise à jour toutes les secondes.</span><span class="sxs-lookup"><span data-stu-id="afa75-199">You should see the current date, which streams an update every second.</span></span> <span data-ttu-id="afa75-200">Bien que cette horloge constitue une seule horloge sur une boucle, vous pouvez utiliser la même idée de définir un minuteur sur des fonctions plus complexes qui effectuent des requêtes Web pour des données en temps réel.</span><span class="sxs-lookup"><span data-stu-id="afa75-200">While this clock is just a timer on a loop, you can use the same idea of setting a timer on more complex functions that make web requests for real-time data.</span></span></li>
</ol>

## <a name="next-steps"></a><span data-ttu-id="afa75-201">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="afa75-201">Next steps</span></span>

<span data-ttu-id="afa75-202">Félicitations !</span><span class="sxs-lookup"><span data-stu-id="afa75-202">Congratulations!</span></span> <span data-ttu-id="afa75-203">Vous avez créé un nouveau projet de fonctions personnalisées, testé une fonction prédéfinie, créé une fonction personnalisée qui demande des données à partir du Web et créé une fonction personnalisée qui diffuse les données.</span><span class="sxs-lookup"><span data-stu-id="afa75-203">You've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams data.</span></span> <span data-ttu-id="afa75-204">Vous pouvez également essayer de déboguer cette fonction à l’aide [des instructions de débogage de la fonction personnalisée](../excel/custom-functions-debugging.md).</span><span class="sxs-lookup"><span data-stu-id="afa75-204">You can also try out debugging this function using [the custom function debugging instructions](../excel/custom-functions-debugging.md).</span></span> <span data-ttu-id="afa75-205">Pour en savoir plus sur les fonctions personnalisées dans Excel, passez à l’article suivant :</span><span class="sxs-lookup"><span data-stu-id="afa75-205">To learn more about custom functions in Excel, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="afa75-206">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="afa75-206">Create custom functions in Excel</span></span>](../excel/custom-functions-overview.md)
