---
title: Didacticiel de fonctions personnalisées Excel
description: Dans ce didacticiel, vous allez créer un complément Excel qui contient une fonction personnalisée qui effectue des calculs, requiert des données web ou lance un flux de données web.
ms.date: 01/16/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 9c8cfedd5f8219f2105456597d43201068b4c21e
ms.sourcegitcommit: a0262ea40cd23f221e69bcb0223110f011265d13
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42688660"
---
# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="767c6-103">Didacticiel : créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="767c6-103">Tutorial: Create custom functions in Excel</span></span>

<span data-ttu-id="767c6-104">Les fonctions personnalisées vous permettent d’ajouter de nouvelles fonctions dans Excel en définissant ces fonctions dans JavaScript dans le cadre d’un complément.</span><span class="sxs-lookup"><span data-stu-id="767c6-104">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="767c6-105">Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe quelle fonction native d’Excel, telle que `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="767c6-105">Users within Excel can access custom functions as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="767c6-106">Vous pouvez créer des fonctions personnalisées qui effectuent des tâches simples comme des calculs personnalisés ou des tâches plus complexes telles que la diffusion en continu des données en temps réel à partir du web dans une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="767c6-106">You can create custom functions that perform simple tasks like calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="767c6-107">Dans ce didacticiel, vous allez :</span><span class="sxs-lookup"><span data-stu-id="767c6-107">In this tutorial, you will:</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="767c6-108">Créer un complément de fonction personnalisée à l’aide la [Générateur Yeoman de compléments Office](https://www.npmjs.com/package/generator-office).</span><span class="sxs-lookup"><span data-stu-id="767c6-108">Create a custom function add-in using the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office).</span></span> 
> * <span data-ttu-id="767c6-109">Utiliser une fonction personnalisée prédéfinie pour effectuer un calcul simple.</span><span class="sxs-lookup"><span data-stu-id="767c6-109">Use a prebuilt custom function to perform a simple calculation.</span></span>
> * <span data-ttu-id="767c6-110">Créer une fonction personnalisée qui demande les données à partir du web.</span><span class="sxs-lookup"><span data-stu-id="767c6-110">Create a custom function that gets data from the web.</span></span>
> * <span data-ttu-id="767c6-111">Créer une fonction personnalisée qui diffuse les données en temps réel à partir du web.</span><span class="sxs-lookup"><span data-stu-id="767c6-111">Create a custom function that streams real-time data from the web.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="767c6-112">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="767c6-112">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* <span data-ttu-id="767c6-113">Excel sur Windows (1904 ou version ultérieure, connecté à un abonnement Office 365) ou sur le web</span><span class="sxs-lookup"><span data-stu-id="767c6-113">Excel on Windows (version 1904 or later, connected to Office 365 subscription) or on the web</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="767c6-114">Créer un projet de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="767c6-114">Create a custom functions project</span></span>

 <span data-ttu-id="767c6-115">Pour commencer, vous devez créer le projet de code pour créer votre complément de fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="767c6-115">To start, you'll create the code project to build your custom function add-in.</span></span> <span data-ttu-id="767c6-116">Le [générateur Yeoman de compléments Office](https://www.npmjs.com/package/generator-office) permettra de configurer votre projet avec certaines fonctions personnalisées prédéfinies que vous pouvez essayer. Si vous avez déjà exécuté le démarrage rapide des fonctions personnalisées et généré un projet, continuez à utiliser ce projet et passez à [cette étape](#create-a-custom-function-that-requests-data-from-the-web).</span><span class="sxs-lookup"><span data-stu-id="767c6-116">The [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) will set up your project with some prebuilt custom functions that you can try out. If you have already run the custom functions quick start and generated a project, continue to use that project and skip to [this step](#create-a-custom-function-that-requests-data-from-the-web) instead.</span></span>

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]
    
    * <span data-ttu-id="767c6-117">**Sélectionnez un type de projet :** `Excel Custom Functions Add-in project`</span><span class="sxs-lookup"><span data-stu-id="767c6-117">**Choose a project type:** `Excel Custom Functions Add-in project`</span></span>
    * <span data-ttu-id="767c6-118">**Sélectionnez un type de script :** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="767c6-118">**Choose a script type:** `JavaScript`</span></span>
    * <span data-ttu-id="767c6-119">**Comment souhaitez-vous nommer votre complément ?**</span><span class="sxs-lookup"><span data-stu-id="767c6-119">**What do you want to name your add-in?**</span></span> `starcount`

    ![Le générateur de yeoman pour les compléments Office vous invite pour les fonctions personnalisées](../images/starcountPrompt.png)
    
    <span data-ttu-id="767c6-121">Le générateur crée le projet et installe les composants Node.js de la prise en charge.</span><span class="sxs-lookup"><span data-stu-id="767c6-121">The Yeoman generator will create the project files and install supporting Node components.</span></span>

    [!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

2. <span data-ttu-id="767c6-122">Accédez au dossier racine du projet.</span><span class="sxs-lookup"><span data-stu-id="767c6-122">Navigate to the root folder of the project.</span></span>
    
    ```command&nbsp;line
    cd starcount
    ```

3. <span data-ttu-id="767c6-123">Créez le projet.</span><span class="sxs-lookup"><span data-stu-id="767c6-123">Build the project.</span></span>
    
    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > <span data-ttu-id="767c6-124">Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez.</span><span class="sxs-lookup"><span data-stu-id="767c6-124">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="767c6-125">Si vous êtes invité à installer un certificat après avoir exécuté `npm run build`, acceptez d’installer le certificat fourni par le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="767c6-125">If you are prompted to install a certificate after you run `npm run build`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

4. <span data-ttu-id="767c6-126">Démarrez le serveur web local qui est exécuté dans Node.js.</span><span class="sxs-lookup"><span data-stu-id="767c6-126">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="767c6-127">Vous pouvez tester le complément de fonction personnalisée dans Excel sur le web ou sur Windows.</span><span class="sxs-lookup"><span data-stu-id="767c6-127">You can try out the custom function add-in in Excel on the web or Windows.</span></span>

# <a name="excel-on-windows-or-mac"></a>[<span data-ttu-id="767c6-128">Excel sur Windows ou Mac</span><span class="sxs-lookup"><span data-stu-id="767c6-128">Excel on Windows or Mac</span></span>](#tab/excel-windows)

<span data-ttu-id="767c6-129">Pour tester votre complément dans Excel sur Windows ou Mac, exécutez la commande suivante.</span><span class="sxs-lookup"><span data-stu-id="767c6-129">To test your add-in in Excel on Windows or Mac, run the following command.</span></span> <span data-ttu-id="767c6-130">Lorsque vous exécutez cette commande, le serveur web local et Excel s’ouvrent avec votre complément chargé.</span><span class="sxs-lookup"><span data-stu-id="767c6-130">When you run this command, the local web server will start and Excel will open with your add-in loaded.</span></span>

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-on-the-web"></a>[<span data-ttu-id="767c6-131">Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="767c6-131">Excel on the web</span></span>](#tab/excel-online)

<span data-ttu-id="767c6-132">Pour tester votre complément dans Excel sur un navigateur, exécutez la commande suivante.</span><span class="sxs-lookup"><span data-stu-id="767c6-132">To test your add-in in Excel on a browser, run the following command.</span></span> <span data-ttu-id="767c6-133">Lorsque vous exécutez cette commande, le serveur web local démarre.</span><span class="sxs-lookup"><span data-stu-id="767c6-133">When you run this command, the local web server will start.</span></span>

```command&nbsp;line
npm run start:web
```

<span data-ttu-id="767c6-134">Pour utiliser votre complément de fonctions personnalisées, ouvrez un nouveau classeur dans Excel sur le web.</span><span class="sxs-lookup"><span data-stu-id="767c6-134">To use your custom functions add-in, open a new workbook in Excel on the web.</span></span> <span data-ttu-id="767c6-135">Dans ce classeur, chargez une version test de votre complément en procédant comme suit.</span><span class="sxs-lookup"><span data-stu-id="767c6-135">In this workbook, complete the following steps to sideload your add-in.</span></span>

1. <span data-ttu-id="767c6-136">Dans Excel, sélectionnez l’onglet **Insertion**, puis **Compléments**.</span><span class="sxs-lookup"><span data-stu-id="767c6-136">In Excel, choose the **Insert** tab and then choose **Add-ins**.</span></span>

   ![Ruban Insertion dans Excel sur le web avec l’icône Mes compléments mise en évidence](../images/excel-cf-online-register-add-in-1.png)
   
2. <span data-ttu-id="767c6-138">Sélectionnez**Gérer mes Compléments** et sélectionnez **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="767c6-138">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="767c6-139">Sélectionnez \*\*Parcourir... \*\* et accédez au répertoire racine du projet créé par le Générateur de Yo Office.</span><span class="sxs-lookup"><span data-stu-id="767c6-139">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="767c6-140">Sélectionnez le fichier**manifest.xml** puis sélectionnez**Ouvrir**, puis sélectionnez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="767c6-140">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 
    
## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="767c6-141">Essayer une fonction personnalisée prédéfinie</span><span class="sxs-lookup"><span data-stu-id="767c6-141">Try out a prebuilt custom function</span></span>

<span data-ttu-id="767c6-142">Le projet de fonctions personnalisées que vous avez créé contient certaines fonctions personnalisées prédéfinies, définies dans le fichier **./src/functions/functions.js**.</span><span class="sxs-lookup"><span data-stu-id="767c6-142">The custom functions project that you created contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file.</span></span> <span data-ttu-id="767c6-143">Le fichier**manifest.xml**indique que toutes les fonctions personnalisées appartiennent à l’`CONTOSO`espace de noms.</span><span class="sxs-lookup"><span data-stu-id="767c6-143">The **./manifest.xml** file specifies that all custom functions belong to the `CONTOSO` namespace.</span></span> <span data-ttu-id="767c6-144">L’espace de noms CONTOSO permet d’accéder aux fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="767c6-144">You'll use the CONTOSO namespace to access the custom functions in Excel.</span></span>

<span data-ttu-id="767c6-145">Essayez de reproduire la`ADD` fonction personnalisée en complétant les étapes suivantes dans Excel:</span><span class="sxs-lookup"><span data-stu-id="767c6-145">Next you'll try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="767c6-146">Dans Excel, accédez à n’importe quelle cellule et entrez `=CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="767c6-146">In Excel, go to any cell and enter `=CONTOSO`.</span></span> <span data-ttu-id="767c6-147">Notez que le menu de saisie semi-automatique affiche la liste de toutes les fonctions dans l’espace de noms `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="767c6-147">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="767c6-148">Exécutez la`CONTOSO.ADD` fonction, avec les nombres `10` et `200` comme paramètres d’entrée, en spécifiant la valeur`=CONTOSO.ADD(10,200)`suivante dans la cellule et appuyez sur entrée.</span><span class="sxs-lookup"><span data-stu-id="767c6-148">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="767c6-149">Le `ADD` fonction personnalisée calcule la somme des deux nombres que vous avez spécifiés et renvoie le résultat**210** .</span><span class="sxs-lookup"><span data-stu-id="767c6-149">The `ADD` custom function computes the sum of the two numbers that you provided and returns the result of **210**.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="767c6-150">Créer une fonction personnalisée qui demande les données à partir du web</span><span class="sxs-lookup"><span data-stu-id="767c6-150">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="767c6-151">Intégration de données à partir du Web est un excellent moyen pour étendre Excel via les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="767c6-151">Integrating data from the Web is a great way to extend Excel through custom functions.</span></span> <span data-ttu-id="767c6-152">Vous allez ensuite créer une fonction personnalisée nommée `getStarCount` qui affiche le nombre d’étoiles attribuées à un référentiel GitHub donné.</span><span class="sxs-lookup"><span data-stu-id="767c6-152">Next you’ll create a custom function named `getStarCount` that shows how many stars a given Github repository possesses.</span></span>

1. <span data-ttu-id="767c6-153">Dans le projet **starcount**, recherchez le fichier **./src/functions/functions.js** et ouvrez-le dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="767c6-153">In the **starcount** project, find the file **./src/functions/functions.js** and open it in your code editor.</span></span> 

2. <span data-ttu-id="767c6-154">Dans **function. js**, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="767c6-154">In **function.js**, add the following code:</span></span> 

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

3. <span data-ttu-id="767c6-155">Exécutez la commande suivante pour régénérer le projet.</span><span class="sxs-lookup"><span data-stu-id="767c6-155">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

4. <span data-ttu-id="767c6-156">Enregistrez de nouveau le complément dans Excel en procédant comme suit (pour Excel sur le web, Windows ou Mac).</span><span class="sxs-lookup"><span data-stu-id="767c6-156">Complete the following steps (for Excel on the web, Windows, or Mac) to re-register the add-in in Excel.</span></span> <span data-ttu-id="767c6-157">Vous devez suivre ces étapes pour que la nouvelle fonction devienne disponible.</span><span class="sxs-lookup"><span data-stu-id="767c6-157">You must complete these steps before the new function will be available.</span></span>

### <a name="excel-on-windows-or-mac"></a>[<span data-ttu-id="767c6-158">Excel sur Windows ou Mac</span><span class="sxs-lookup"><span data-stu-id="767c6-158">Excel on Windows or Mac</span></span>](#tab/excel-windows)

1. <span data-ttu-id="767c6-159">Fermez Excel, puis rouvrez-le.</span><span class="sxs-lookup"><span data-stu-id="767c6-159">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="767c6-160">Dans Excel, sélectionnez l’onglet **Insertion**, puis cliquez sur la flèche vers le bas située à droite de **Mes compléments**.  ![Ruban Insertion dans Excel pour Windows avec la flèche Mes compléments mise en évidence](../images/select-insert.png)</span><span class="sxs-lookup"><span data-stu-id="767c6-160">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel on Windows with the My Add-ins arrow highlighted](../images/select-insert.png)</span></span>

3. <span data-ttu-id="767c6-161">Dans la liste des compléments disponibles, recherchez la section **Compléments de développeur**, puis sélectionnez le complément **starcount** pour effectuer cette opération.</span><span class="sxs-lookup"><span data-stu-id="767c6-161">In the list of available add-ins, find the **Developer Add-ins** section and select the **starcount** add-in to register it.</span></span>
    <span data-ttu-id="767c6-162">![Ruban Insertion dans Excel sur Windows avec le complément Fonctions personnalisées Excel mis en évidence dans la liste Mes compléments](../images/list-starcount.png)</span><span class="sxs-lookup"><span data-stu-id="767c6-162">![Insert ribbon in Excel on Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/list-starcount.png)</span></span>


# <a name="excel-on-the-web"></a>[<span data-ttu-id="767c6-163">Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="767c6-163">Excel on the web</span></span>](#tab/excel-online)

1. <span data-ttu-id="767c6-164">Dans Excel, sélectionnez l’onglet **Insertion**, puis **Compléments**.  ![Ruban Insertion dans Excel sur le web avec l’icône Mes compléments mise en évidence](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="767c6-164">In Excel, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel on the web with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="767c6-165">Sélectionnez**Gérer mes Compléments** et sélectionnez **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="767c6-165">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="767c6-166">Sélectionnez \*\*Parcourir... \*\* et accédez au répertoire racine du projet créé par le Générateur de Yo Office.</span><span class="sxs-lookup"><span data-stu-id="767c6-166">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="767c6-167">Sélectionnez le fichier**manifest.xml** puis sélectionnez**Ouvrir**, puis sélectionnez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="767c6-167">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

<ol start="5">
<li> <span data-ttu-id="767c6-168">Essayez la nouvelle fonction.</span><span class="sxs-lookup"><span data-stu-id="767c6-168">Try out the new function.</span></span> <span data-ttu-id="767c6-169">Dans la cellule <strong>B1</strong>, tapez le texte <strong>=CONTOSO.GETSTARCOUNT("OfficeDev", "Excel-Custom-Functions")</strong>, puis appuyez sur Entrée.</span><span class="sxs-lookup"><span data-stu-id="767c6-169">In cell <strong>B1</strong>, type the text <strong>=CONTOSO.GETSTARCOUNT("OfficeDev", "Excel-Custom-Functions")</strong> and press enter.</span></span> <span data-ttu-id="767c6-170">Le résultat dans la cellule <strong>B1</strong> doit correspondre au nombre d’étoiles actuellement attribuées au [référentiel GitHub Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions).</span><span class="sxs-lookup"><span data-stu-id="767c6-170">You should see that the result in cell <strong>B1</strong> is the current number of stars given to the [Excel-Custom-Functions Github repository](https://github.com/OfficeDev/Excel-Custom-Functions).</span></span></li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="767c6-171">Créer une fonction personnalisée asynchrone de diffusion en continu</span><span class="sxs-lookup"><span data-stu-id="767c6-171">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="767c6-172">La fonction `getStarCount` renvoie le nombre d’étoiles attribuées à un référentiel à un moment donné.</span><span class="sxs-lookup"><span data-stu-id="767c6-172">The `getStarCount` function returns the number of stars a repository has at a specific moment in time.</span></span> <span data-ttu-id="767c6-173">Les fonctions personnalisées peuvent également renvoyer des données qui changent continuellement.</span><span class="sxs-lookup"><span data-stu-id="767c6-173">Custom functions can also return data that is continuously changing.</span></span> <span data-ttu-id="767c6-174">Ces fonctions sont appelées fonctions de diffusion en continu.</span><span class="sxs-lookup"><span data-stu-id="767c6-174">These functions are called streaming functions.</span></span> <span data-ttu-id="767c6-175">Elles doivent inclure un paramètre `invocation` qui fait référence à la cellule à partir de laquelle la fonction a été appelée.</span><span class="sxs-lookup"><span data-stu-id="767c6-175">They must include an `invocation` parameter which refers to the cell where the function was called from.</span></span> <span data-ttu-id="767c6-176">Le paramètre `invocation` permet de mettre à jour le contenu de la cellule à tout moment.</span><span class="sxs-lookup"><span data-stu-id="767c6-176">The `invocation` parameter is used to update the contents of the cell at any time.</span></span>  

<span data-ttu-id="767c6-177">Vous remarquerez que l’exemple de code suivant inclut deux fonctions (`currentTime` et `clock`).</span><span class="sxs-lookup"><span data-stu-id="767c6-177">In the following code sample, you'll notice that there are two functions, `currentTime` and `clock`.</span></span> <span data-ttu-id="767c6-178">`currentTime` est une fonction statique qui n’utilise pas la diffusion en continu.</span><span class="sxs-lookup"><span data-stu-id="767c6-178">The `currentTime` function is a static function that does not use streaming.</span></span> <span data-ttu-id="767c6-179">Elle renvoie la date sous la forme d’une chaîne.</span><span class="sxs-lookup"><span data-stu-id="767c6-179">It returns the date as a string.</span></span> <span data-ttu-id="767c6-180">La fonction `clock` utilise la fonction `currentTime` pour fournir la nouvelle heure toutes les secondes à une cellule dans Excel.</span><span class="sxs-lookup"><span data-stu-id="767c6-180">The `clock` function uses the `currentTime` function to provide the new time every second to a cell in Excel.</span></span> <span data-ttu-id="767c6-181">Elle utilise `invocation.setResult` pour communiquer l’heure à la cellule Excel et `invocation.onCanceled` pour gérer le résultat de l’annulation de la fonction.</span><span class="sxs-lookup"><span data-stu-id="767c6-181">It uses `invocation.setResult` to deliver the time to the Excel cell and `invocation.onCanceled` to handle what occurs when the function is canceled.</span></span>

1. <span data-ttu-id="767c6-182">Dans le projet **starcount**, ajoutez le code suivant à **./src/functions/functions.js**, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="767c6-182">In the **starcount** project, add the following code to **./src/functions/functions.js** and save the file.</span></span>

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

2. <span data-ttu-id="767c6-183">Exécutez la commande suivante pour régénérer le projet.</span><span class="sxs-lookup"><span data-stu-id="767c6-183">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

3. <span data-ttu-id="767c6-184">Enregistrez de nouveau le complément dans Excel en procédant comme suit (pour Excel sur le web, Windows ou Mac).</span><span class="sxs-lookup"><span data-stu-id="767c6-184">Complete the following steps (for Excel on the web, Windows, or Mac) to re-register the add-in in Excel.</span></span> <span data-ttu-id="767c6-185">Vous devez suivre ces étapes pour que la nouvelle fonction devienne disponible.</span><span class="sxs-lookup"><span data-stu-id="767c6-185">You must complete these steps before the new function will be available.</span></span> 

# <a name="excel-on-windows-or-mac"></a>[<span data-ttu-id="767c6-186">Excel sur Windows ou Mac</span><span class="sxs-lookup"><span data-stu-id="767c6-186">Excel on Windows or Mac</span></span>](#tab/excel-windows)

1. <span data-ttu-id="767c6-187">Fermez Excel, puis rouvrez-le.</span><span class="sxs-lookup"><span data-stu-id="767c6-187">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="767c6-188">Dans Excel, sélectionnez l’onglet **Insertion**, puis cliquez sur la flèche vers le bas située à droite de **Mes compléments**.  ![Ruban Insertion dans Excel pour Windows avec la flèche Mes compléments mise en évidence](../images/select-insert.png)</span><span class="sxs-lookup"><span data-stu-id="767c6-188">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel on Windows with the My Add-ins arrow highlighted](../images/select-insert.png)</span></span>

3. <span data-ttu-id="767c6-189">Dans la liste des compléments disponibles, recherchez la section **Compléments de développeur**, puis sélectionnez le complément **starcount** pour effectuer cette opération.</span><span class="sxs-lookup"><span data-stu-id="767c6-189">In the list of available add-ins, find the **Developer Add-ins** section and select the **starcount** add-in to register it.</span></span>
    <span data-ttu-id="767c6-190">![Ruban Insertion dans Excel sur Windows avec le complément Fonctions personnalisées Excel mis en évidence dans la liste Mes compléments](../images/list-starcount.png)</span><span class="sxs-lookup"><span data-stu-id="767c6-190">![Insert ribbon in Excel on Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/list-starcount.png)</span></span>

# <a name="excel-on-the-web"></a>[<span data-ttu-id="767c6-191">Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="767c6-191">Excel on the web</span></span>](#tab/excel-online)

1. <span data-ttu-id="767c6-192">Dans Excel, sélectionnez l’onglet **Insertion**, puis **Compléments**.  ![Ruban Insertion dans Excel sur le web avec l’icône Mes compléments mise en évidence](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="767c6-192">In Excel, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel on the web with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="767c6-193">Sélectionnez**Gérer mes Compléments** et sélectionnez **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="767c6-193">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="767c6-194">Sélectionnez \*\*Parcourir... \*\* et accédez au répertoire racine du projet créé par le Générateur de Yo Office.</span><span class="sxs-lookup"><span data-stu-id="767c6-194">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="767c6-195">Sélectionnez le fichier**manifest.xml** puis sélectionnez**Ouvrir**, puis sélectionnez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="767c6-195">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 

<ol start="4">
<li><span data-ttu-id="767c6-196">Essayez la nouvelle fonction.</span><span class="sxs-lookup"><span data-stu-id="767c6-196">Try out the new function.</span></span> <span data-ttu-id="767c6-197">Dans la cellule <strong>C1</strong>, tapez le texte <strong>=CONTOSO.CLOCK())</strong>, puis appuyez sur Entrée.</span><span class="sxs-lookup"><span data-stu-id="767c6-197">In cell <strong>C1</strong>, type the text <strong>=CONTOSO.CLOCK())</strong> and press enter.</span></span> <span data-ttu-id="767c6-198">La date du jour doit apparaître. Elle est mise à jour toutes les secondes.</span><span class="sxs-lookup"><span data-stu-id="767c6-198">You should see the current date, which streams an update every second.</span></span> <span data-ttu-id="767c6-199">Cette horloge n’est qu’une minuterie incluse dans une boucle, mais vous pouvez vous inspirer de cette idée pour créer des fonctions plus complexes qui récupèrent des données en temps réel en exécutant des requêtes web.</span><span class="sxs-lookup"><span data-stu-id="767c6-199">While this clock is just a timer on a loop, you can use the same idea of setting a timer on more complex functions that make web requests for real-time data.</span></span></li>
</ol>

## <a name="next-steps"></a><span data-ttu-id="767c6-200">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="767c6-200">Next steps</span></span>

<span data-ttu-id="767c6-201">Félicitations !</span><span class="sxs-lookup"><span data-stu-id="767c6-201">Congratulations!</span></span> <span data-ttu-id="767c6-202">Vous avez créé un nouveau projet de fonctions personnalisées, essayé une fonction prédéfinie, créé une fonction personnalisée qui récupère des données à partir du web et créé une fonction personnalisée qui diffuse des données.</span><span class="sxs-lookup"><span data-stu-id="767c6-202">You've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams data.</span></span> <span data-ttu-id="767c6-203">Vous pouvez également essayer de déboguer cette fonction à l’aide des [instructions de débogage des fonction personnalisées](../excel/custom-functions-debugging.md).</span><span class="sxs-lookup"><span data-stu-id="767c6-203">You can also try out debugging this function using [the custom function debugging instructions](../excel/custom-functions-debugging.md).</span></span> <span data-ttu-id="767c6-204">Pour en savoir plus sur les fonctions personnalisées dans Excel, passez à l’article suivant :</span><span class="sxs-lookup"><span data-stu-id="767c6-204">To learn more about custom functions in Excel, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="767c6-205">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="767c6-205">Create custom functions in Excel</span></span>](../excel/custom-functions-overview.md)
