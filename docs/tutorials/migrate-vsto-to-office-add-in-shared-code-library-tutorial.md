---
ms.date: 02/09/2021
ms.prod: non-product-specific
description: Didacticiel sur le partage de codes entre un complément VSTO et un complément Office.
title: 'Didacticiel : partage de codes entre un complément VSTO et un complément Office à l’aide d’une bibliothèque de codes partagée'
localization_priority: Priority
ms.openlocfilehash: 1645cdcc3c799ec09e98ae69dd4abd6e38b11880
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50238091"
---
# <a name="tutorial-share-code-between-both-a-vsto-add-in-and-an-office-add-in-with-a-shared-code-library"></a><span data-ttu-id="cbb82-103">Didacticiel : partage de codes entre un complément VSTO et un complément Office avec une bibliothèque de codes partagée</span><span class="sxs-lookup"><span data-stu-id="cbb82-103">Tutorial: Share code between both a VSTO Add-in and an Office Add-in with a shared code library</span></span>

<span data-ttu-id="cbb82-104">Les compléments de Visual Studio Tools pour Office (VSTO) sont idéaux pour étendre Office afin de fournir des solutions aux entreprises, la vôtre ou d’autres.</span><span class="sxs-lookup"><span data-stu-id="cbb82-104">Visual Studio Tools for Office (VSTO) Add-ins are great for extending Office to provide solutions for your business or others.</span></span> <span data-ttu-id="cbb82-105">Ils existent depuis longtemps et des milliers de solutions sont créées avec VSTO.</span><span class="sxs-lookup"><span data-stu-id="cbb82-105">They've been around for a long time and there are thousands of solutions built with VSTO.</span></span> <span data-ttu-id="cbb82-106">Cependant, ils s’exécutent uniquement avec Office sur Windows.</span><span class="sxs-lookup"><span data-stu-id="cbb82-106">However, they only run on Office on Windows.</span></span> <span data-ttu-id="cbb82-107">Vous ne pouvez pas exécuter des compléments VSTO sur les plateformes Mac, Online ou mobile.</span><span class="sxs-lookup"><span data-stu-id="cbb82-107">You can't run VSTO Add-ins on Mac, online, or mobile platforms.</span></span>

<span data-ttu-id="cbb82-108">Les compléments Office utilisent HTML, JavaScript et d’autres technologies web pour créer des solutions Office sur toutes les plateformes.</span><span class="sxs-lookup"><span data-stu-id="cbb82-108">Office Add-ins use HTML, JavaScript, and additional web technologies to build Office solutions on all platforms.</span></span> <span data-ttu-id="cbb82-109">La migration de votre complément VSTO existant vers un complément Office est un excellent moyen de rendre votre solution accessible sur toutes les plateformes.</span><span class="sxs-lookup"><span data-stu-id="cbb82-109">Migrating your existing VSTO Add-in to an Office Add-in is a great way to make your solution available across all platforms.</span></span>

<span data-ttu-id="cbb82-110">Vous pouvez conserver votre complément VSTO et un nouveau complément Office ayant les mêmes fonctionnalités.</span><span class="sxs-lookup"><span data-stu-id="cbb82-110">You may want to maintain both your VSTO Add-in and a new Office Add-in that both have the same functionality.</span></span> <span data-ttu-id="cbb82-111">Cela vous permet de continuer à offrir un service à vos clients qui utilisent le complément VSTO pour Office sur Windows.</span><span class="sxs-lookup"><span data-stu-id="cbb82-111">This enables you to continue servicing your customers that use the VSTO Add-in on Office on Windows.</span></span> <span data-ttu-id="cbb82-112">Cela vous permet également de proposer aux clients la même fonctionnalité dans un complément Office pour l'ensemble des plateformes.</span><span class="sxs-lookup"><span data-stu-id="cbb82-112">This also enables you to provide the same functionality in an Office Add-in for customers across all platforms.</span></span> <span data-ttu-id="cbb82-113">Vous pouvez également [Rendre votre complément Office compatible avec le complément VSTO existant](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="cbb82-113">You can also [Make your Office Add-in compatible with the existing VSTO Add-in](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).</span></span>

<span data-ttu-id="cbb82-114">Il est toutefois préférable d’éviter la réécriture du code entier de votre complément VSTO pour le complément Office.</span><span class="sxs-lookup"><span data-stu-id="cbb82-114">However it is best to avoid rewriting all the code from your VSTO Add-in for the Office Add-in.</span></span> <span data-ttu-id="cbb82-115">Ce didacticiel explique les précautions à prendre pour éviter la réécriture d'un code grâce à l'utilisation d’une bibliothèque de codes partagés pour les deux compléments.</span><span class="sxs-lookup"><span data-stu-id="cbb82-115">This tutorial shows how to avoid rewriting code by using a shared code library for both add-ins.</span></span>

## <a name="shared-code-library"></a><span data-ttu-id="cbb82-116">Bibliothèque de codes partagés</span><span class="sxs-lookup"><span data-stu-id="cbb82-116">Shared code library</span></span>

<span data-ttu-id="cbb82-117">Ce didacticiel vous guide dans la procédure d’identification et de partage d'un code commun à votre complément VSTO et à un complément Office moderne.</span><span class="sxs-lookup"><span data-stu-id="cbb82-117">This tutorial will walk you through the steps of identifying and sharing common code between your VSTO Add-in and a modern Office Add-in.</span></span> <span data-ttu-id="cbb82-118">Ce guide utilise un exemple de complément VSTO très simple pour suivre les étapes afin que vous puissiez vous concentrer sur les compétences et les techniques dont vous aurez besoin pour utiliser vos propres compléments VSTO.</span><span class="sxs-lookup"><span data-stu-id="cbb82-118">It uses a very simple VSTO Add-in example for the steps so that you can focus on the skills and techniques you will need for working with your own VSTO Add-ins.</span></span>

<span data-ttu-id="cbb82-119">Le diagramme suivant illustre le fonctionnement de la bibliothèque de codes partagés pour la migration.</span><span class="sxs-lookup"><span data-stu-id="cbb82-119">The following diagram shows how the shared code library works for migration.</span></span> <span data-ttu-id="cbb82-120">Le code commun est refactorisé dans une nouvelle bibliothèque de codes partagés.</span><span class="sxs-lookup"><span data-stu-id="cbb82-120">Common code is refactored into a new shared code library.</span></span> <span data-ttu-id="cbb82-121">Le code peut demeurer écrit dans son langage d’origine, par exemple C# ou VB.</span><span class="sxs-lookup"><span data-stu-id="cbb82-121">The code can remain written in its original language, such as C# or VB.</span></span> <span data-ttu-id="cbb82-122">Cela signifie que vous continuez à utiliser le code dans le complément VSTO existant en créant une référence de projet.</span><span class="sxs-lookup"><span data-stu-id="cbb82-122">This means you can continue using the code in the existing VSTO Add-in by creating a project reference.</span></span> <span data-ttu-id="cbb82-123">Lorsque vous créez le complément Office, celui-ci utilise également la bibliothèque de codes partagés en y appelant les API REST.</span><span class="sxs-lookup"><span data-stu-id="cbb82-123">When you create the Office Add-in, it will also use the shared code library by calling into it through REST APIs.</span></span>

![Diagramme d'un complément VSTO et d'un complément Office utilisant une bibliothèque de codes partagés](../images/vsto-migration-shared-code-library.png)

<span data-ttu-id="cbb82-125">Compétences et techniques décrites dans ce didacticiel :</span><span class="sxs-lookup"><span data-stu-id="cbb82-125">Skills and techniques in this tutorial:</span></span>

- <span data-ttu-id="cbb82-126">Créer une bibliothèque de classes partagées en refactorisant le code dans une bibliothèque de classes .NET.</span><span class="sxs-lookup"><span data-stu-id="cbb82-126">Create a shared class library by refactoring code into a .NET class library.</span></span>
- <span data-ttu-id="cbb82-127">Créez un wrapper API REST à l’aide de ASP.NET Core pour la bibliothèque de classes partagées.</span><span class="sxs-lookup"><span data-stu-id="cbb82-127">Create a REST API wrapper using ASP.NET Core for the shared class library.</span></span>
- <span data-ttu-id="cbb82-128">Appelez l’API REST à partir du complément Office pour accéder au code partagé.</span><span class="sxs-lookup"><span data-stu-id="cbb82-128">Call the REST API from the Office Add-in to access shared code.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="cbb82-129">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="cbb82-129">Prerequisites</span></span>

<span data-ttu-id="cbb82-130">Pour la configuration de votre environnement de développement :</span><span class="sxs-lookup"><span data-stu-id="cbb82-130">To set up your development environment:</span></span>

1. <span data-ttu-id="cbb82-131">Installez [Visual Studio 2019](https://visualstudio.microsoft.com/downloads/).</span><span class="sxs-lookup"><span data-stu-id="cbb82-131">Install [Visual Studio 2019](https://visualstudio.microsoft.com/downloads/).</span></span>
2. <span data-ttu-id="cbb82-132">Installez les charges de travail suivantes :</span><span class="sxs-lookup"><span data-stu-id="cbb82-132">Install the following workloads:</span></span>
    - <span data-ttu-id="cbb82-133">ASP.NET et le développement web</span><span class="sxs-lookup"><span data-stu-id="cbb82-133">ASP.NET and web development</span></span>
    - <span data-ttu-id="cbb82-134">Développement multiplateforme .NET Core.</span><span class="sxs-lookup"><span data-stu-id="cbb82-134">.NET Core cross-platform development.</span></span>
    - <span data-ttu-id="cbb82-135">Développement Office/SharePoint</span><span class="sxs-lookup"><span data-stu-id="cbb82-135">Office/SharePoint development</span></span>
    - <span data-ttu-id="cbb82-136">Les éléments **Individuels** suivants.</span><span class="sxs-lookup"><span data-stu-id="cbb82-136">The following **Individual** components.</span></span>
        - <span data-ttu-id="cbb82-137">Visual Studio Tools pour Office (VSTO).</span><span class="sxs-lookup"><span data-stu-id="cbb82-137">Visual Studio Tools for Office (VSTO).</span></span>
        - <span data-ttu-id="cbb82-138">.NET Core 3.0 Runtime.</span><span class="sxs-lookup"><span data-stu-id="cbb82-138">.NET Core 3.0 Runtime.</span></span>

<span data-ttu-id="cbb82-139">Vous devez également disposer des éléments ci-après :</span><span class="sxs-lookup"><span data-stu-id="cbb82-139">You also need the following:</span></span>

- <span data-ttu-id="cbb82-140">Un compte Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="cbb82-140">A Microsoft 365 account.</span></span> <span data-ttu-id="cbb82-141">Vous pouvez rejoindre le [programme pour les développeurs Microsoft 365](https://aka.ms/devprogramsignup) qui offre un abonnement Microsoft 365 renouvelable de 90 jours qui inclut les applications Office.</span><span class="sxs-lookup"><span data-stu-id="cbb82-141">You can join the [Microsoft 365 developer program](https://aka.ms/devprogramsignup) that provides a renewable 90-day Microsoft 365 subscription that includes Office apps.</span></span>
- <span data-ttu-id="cbb82-142">Un locataire Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="cbb82-142">A Microsoft Azure Tenant.</span></span> <span data-ttu-id="cbb82-143">Un abonnement d’évaluation peut être obtenu ici : [Microsoft Azure](https://account.windowsazure.com/SignUp).</span><span class="sxs-lookup"><span data-stu-id="cbb82-143">A trial subscription can be acquired here: [Microsoft Azure](https://account.windowsazure.com/SignUp).</span></span>

## <a name="the-cell-analyzer-vsto-add-in"></a><span data-ttu-id="cbb82-144">Le composant VSTO d’analyseur de cellule</span><span class="sxs-lookup"><span data-stu-id="cbb82-144">The Cell analyzer VSTO Add-in</span></span>

<span data-ttu-id="cbb82-145">Ce didacticiel utilise la solution PnP pour la [Bibliothèque de compléments VSTO partagés pour les compléments Office](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/VSTO-shared-code-migration).</span><span class="sxs-lookup"><span data-stu-id="cbb82-145">This tutorial uses the [VSTO Add-in shared library for Office Add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/VSTO-shared-code-migration) PnP solution.</span></span> <span data-ttu-id="cbb82-146">Le dossier **/Start** contient la solution de complément VSTO que vous allez migrer.</span><span class="sxs-lookup"><span data-stu-id="cbb82-146">The **/start** folder contains the VSTO Add-in solution that you will migrate.</span></span> <span data-ttu-id="cbb82-147">Votre objectif est de migrer le complément VSTO vers un complément Office moderne en partageant le code lorsque cela est possible.</span><span class="sxs-lookup"><span data-stu-id="cbb82-147">Your goal is to migrate the VSTO Add-in to a modern Office Add-in by sharing code when possible.</span></span>

> [!NOTE]
> <span data-ttu-id="cbb82-148">L’exemple utilise C# , mais vous pouvez utiliser les techniques décrites dans ce didacticiel pour appliquer un complément VSTO écrit dans n’importe quel langage .NET.</span><span class="sxs-lookup"><span data-stu-id="cbb82-148">The sample uses C# but you can apply the techniques in this tutorial to a VSTO Add-in written in any .NET language.</span></span>

1. <span data-ttu-id="cbb82-149">Téléchargez la solution PnP pour la [Bibliothèque de compléments VSTO partagés pour les compléments Office](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/VSTO-shared-code-migration) vers un dossier de travail de votre ordinateur.</span><span class="sxs-lookup"><span data-stu-id="cbb82-149">Download the [VSTO Add-in shared library for Office Add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/VSTO-shared-code-migration) PnP solution to a working folder on your computer.</span></span>
1. <span data-ttu-id="cbb82-150">Démarrez Visual Studio 2019 et ouvrez la solution **/start/Cell-Analyzer.sln**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-150">Start Visual Studio 2019 and open the **/start/Cell-Analyzer.sln** solution.</span></span>
1. <span data-ttu-id="cbb82-151">Dans le menu **Déboguer**, choisissez **Démarrer le débogage**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-151">On the **Debug** menu, choose **Start Debugging**.</span></span>
1. <span data-ttu-id="cbb82-152">Dans l’**Explorateur de solutions**, cliquez à l'aide du bouton droit sur le projet **Analyseur de cellule**, puis choisissez **Propriétés**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-152">In **Solution Explorer**, right-click the **Cell-Analyzer** project, and choose **Properties**.</span></span>
1. <span data-ttu-id="cbb82-153">Sélectionnez la catégorie de **Signature** dans les propriétés.</span><span class="sxs-lookup"><span data-stu-id="cbb82-153">Choose the **Signing** category in the properties.</span></span>
1. <span data-ttu-id="cbb82-154">Sélectionnez **Signer des manifestes ClickOnce**, puis choisissez **Créer un certificat de test**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-154">Choose **Sign the ClickOnce manifests**, and then chose **Create Test Certificate**.</span></span>
1. <span data-ttu-id="cbb82-155">Dans la boîte de dialogue **Créer un certificat de test**, entrez et confirmez un mot de passe.</span><span class="sxs-lookup"><span data-stu-id="cbb82-155">In the **Create Test Certificate** dialog, enter and confirm a password.</span></span> <span data-ttu-id="cbb82-156">Sélectionnez ensuite **OK**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-156">Then choose **OK**.</span></span>

<span data-ttu-id="cbb82-157">Le complément est un volet de tâche personnalisé Office pour Excel.</span><span class="sxs-lookup"><span data-stu-id="cbb82-157">The add-in is a custom task pane for Excel.</span></span> <span data-ttu-id="cbb82-158">Vous pouvez sélectionner n’importe quelle cellule contenant un texte, puis choisissez le bouton **Afficher les Unicodes**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-158">You can select any cell with text, and then choose the **Show unicode** button.</span></span> <span data-ttu-id="cbb82-159">Dans la section **Résultat** , le complément affiche une liste de chaque caractère du texte, ainsi que leur nombre Unicode correspondant.</span><span class="sxs-lookup"><span data-stu-id="cbb82-159">In the **Result** section, the add-in will display a list of each character in the text along with its corresponding Unicode number.</span></span>

![Capture d’écran du complément VSTO d’analyseur de cellule exécuté dans Excel avec le bouton Afficher Unicode et la section Résultat vide](../images/pnp-cell-analyzer-vsto-add-in.png)

## <a name="analyze-types-of-code-in-the-vsto-add-in"></a><span data-ttu-id="cbb82-161">Analyser les types de code dans le complément VSTO</span><span class="sxs-lookup"><span data-stu-id="cbb82-161">Analyze types of code in the VSTO Add-in</span></span>

<span data-ttu-id="cbb82-162">La première technique à appliquer consiste à analyser le complément pour identifier les parties de code pouvant être partagées.</span><span class="sxs-lookup"><span data-stu-id="cbb82-162">The first technique to apply is to analyze the add-in for which parts of code can be shared.</span></span> <span data-ttu-id="cbb82-163">Un projet se décompose généralement en trois types de codes.</span><span class="sxs-lookup"><span data-stu-id="cbb82-163">In general, project will break down into three types of code.</span></span>

### <a name="ui-code"></a><span data-ttu-id="cbb82-164">Code d'interface utilisateur</span><span class="sxs-lookup"><span data-stu-id="cbb82-164">UI code</span></span>

<span data-ttu-id="cbb82-165">Le code d'interface utilisateur communique avec l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="cbb82-165">UI code interacts with the user.</span></span> <span data-ttu-id="cbb82-166">Dans VSTO, le code d'interface utilisation fonctionne par le biais de Windows Forms.</span><span class="sxs-lookup"><span data-stu-id="cbb82-166">In VSTO UI code works through Windows Forms.</span></span> <span data-ttu-id="cbb82-167">Les compléments Office utilisent les langages HTML, CSS et JavaScript pour l'interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="cbb82-167">Office Add-ins use HTML, CSS, and JavaScript for UI.</span></span> <span data-ttu-id="cbb82-168">Vous ne pouvez pas, en raison de ces différences, partager le code d’interface utilisateur avec le complément Office.</span><span class="sxs-lookup"><span data-stu-id="cbb82-168">Because of these differences you cannot share UI code to the Office Add-in.</span></span> <span data-ttu-id="cbb82-169">L’interface utilisateur doit être recréé dans JavaScript.</span><span class="sxs-lookup"><span data-stu-id="cbb82-169">UI will need to be recreated in JavaScript.</span></span>

### <a name="document-code"></a><span data-ttu-id="cbb82-170">Code de document</span><span class="sxs-lookup"><span data-stu-id="cbb82-170">Document code</span></span>

<span data-ttu-id="cbb82-171">Le code communique avec le document par le biais d’objets .NET tels que `Microsoft.Office.Interop.Excel.Range` dans VSTO.</span><span class="sxs-lookup"><span data-stu-id="cbb82-171">In VSTO code interacts with the document through .NET objects such as `Microsoft.Office.Interop.Excel.Range`.</span></span> <span data-ttu-id="cbb82-172">Les compléments Office utilisent néanmoins la bibliothèque Office.js.</span><span class="sxs-lookup"><span data-stu-id="cbb82-172">But Office Add-ins use the Office.js library.</span></span> <span data-ttu-id="cbb82-173">Ils ne sont pas exactement identiques, bien qu'ils soient similaires.</span><span class="sxs-lookup"><span data-stu-id="cbb82-173">Although these are similar, they are not exactly the same.</span></span> <span data-ttu-id="cbb82-174">Par conséquent, vous ne pouvez pas partager le code d'interaction d'un document avec le complément Office.</span><span class="sxs-lookup"><span data-stu-id="cbb82-174">So again, you cannot share document interaction code to the Office Add-in.</span></span>

### <a name="logic-code"></a><span data-ttu-id="cbb82-175">Code logique</span><span class="sxs-lookup"><span data-stu-id="cbb82-175">Logic code</span></span>

<span data-ttu-id="cbb82-176">La logique métier, les algorithmes, les fonctions d’assistance et autres codes similaires constituent souvent le cœur d’un complément VSTO.</span><span class="sxs-lookup"><span data-stu-id="cbb82-176">Business logic, algorithms, helper functions, and similar code often make up the heart of a VSTO Add-in.</span></span> <span data-ttu-id="cbb82-177">Ce code fonctionne indépendamment de l’interface utilisateur et du code de document pour effectuer une analyse, se connecter à un service principale, effectuer des calculs, etc.</span><span class="sxs-lookup"><span data-stu-id="cbb82-177">This code works independently of the UI and document code to perform analysis, connect to backend services, run calculations, and more.</span></span> <span data-ttu-id="cbb82-178">Il s’agit du code qui peut être partagé pour que vous n’ayez pas à le réécrire dans JavaScript.</span><span class="sxs-lookup"><span data-stu-id="cbb82-178">This is the code that can be shared so that you don't have to rewrite it in JavaScript.</span></span>

<span data-ttu-id="cbb82-179">Examinez le complément VSTO.</span><span class="sxs-lookup"><span data-stu-id="cbb82-179">Let's examine the VSTO Add-in.</span></span> <span data-ttu-id="cbb82-180">Dans le code suivant, chaque section est identifiée en tant que code de DOCUMENT, d’interface utilisateur ou d’ALGORITHME.</span><span class="sxs-lookup"><span data-stu-id="cbb82-180">In the following code, each section is identified as DOCUMENT, UI, or ALGORITHM code.</span></span>

```csharp
// *** UI CODE ***
private void btnUnicode_Click(object sender, EventArgs e)
{
    // *** DOCUMENT CODE ***
    Microsoft.Office.Interop.Excel.Range rangeCell;
    rangeCell = Globals.ThisAddIn.Application.ActiveCell;

    string cellValue = "";

    if (null != rangeCell.Value)
    {
        cellValue = rangeCell.Value.ToString();
    }

    // *** ALGORITHM CODE ***
    //convert string to Unicode listing
    string result = "";
    foreach (char c in cellValue)
    {
        int unicode = c;

        result += $"{c}: {unicode}\r\n";
    }

    // *** UI CODE ***
    //Output the result
    txtResult.Text = result;
}
```

<span data-ttu-id="cbb82-181">Grâce à cette approche, vous pouvez voir qu’une section de code peut être partagée avec le complément Office.</span><span class="sxs-lookup"><span data-stu-id="cbb82-181">Using this approach you can see that one section of code can be shared to the Office Add-in.</span></span> <span data-ttu-id="cbb82-182">Le code suivant doit être refactorisé dans une bibliothèque de classes distincte.</span><span class="sxs-lookup"><span data-stu-id="cbb82-182">The following code will need to be refactored into a separate class library.</span></span>

```csharp
// *** ALGORITHM CODE ***
//convert string to Unicode listing
string result = "";
foreach (char c in cellValue)
{
    int unicode = c;

    result += $"{c}: {unicode}\r\n";
}
```

## <a name="create-a-shared-class-library"></a><span data-ttu-id="cbb82-183">Créer une bibliothèque de classes partagées</span><span class="sxs-lookup"><span data-stu-id="cbb82-183">Create a shared class library</span></span>

<span data-ttu-id="cbb82-184">Les compléments VSTO étant créés dans Visual Studio en tant que projets .NET, nous réutiliser .NET aussi souvent que possible pour simplifier les choses.</span><span class="sxs-lookup"><span data-stu-id="cbb82-184">VSTO Add-ins are created in Visual Studio as .NET projects, so we'll reuse .NET as much as possible to keep things simple.</span></span> <span data-ttu-id="cbb82-185">La technique suivante consiste à créer une bibliothèque de classes et à refactoriser le code partagé dans cette bibliothèque.</span><span class="sxs-lookup"><span data-stu-id="cbb82-185">Our next technique is to create a class library and refactor shared code into that class library.</span></span>

1. <span data-ttu-id="cbb82-186">Si ce n'est pas encore fait, démarrez Visual Studio 2019 et ouvrez la solution **/start/Cell-Analyzer.sln**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-186">If you haven't already, start Visual Studio 2019 and open the **\start\Cell-Analyzer.sln** solution.</span></span>
2. <span data-ttu-id="cbb82-187">Cliquez avec le bouton droit sur la solution dans l’**Explorateur de solutions** et choisissez **Ajouter > Nouvelle solution**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-187">Right-click the solution in **Solution Explorer** and choose **Add > New Project**.</span></span>
3. <span data-ttu-id="cbb82-188">Dans la **boîte de dialogue Ajouter un nouveau projet**, choisissez **Bibliothèque de classes (.NET Framework)**, puis sélectionnez **Suivant**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-188">In the **Add a new project dialog**, choose **Class Library (.NET Framework)**, and choose **Next**.</span></span>
    > [!NOTE]
    > <span data-ttu-id="cbb82-189">N’utilisez pas la bibliothèque de classes .NET Core, car elle ne fonctionnera pas avec votre projet VSTO.</span><span class="sxs-lookup"><span data-stu-id="cbb82-189">Don't use the .NET Core class library because it will not work with your VSTO project.</span></span>
4. <span data-ttu-id="cbb82-190">Dans la boîte de dialogue **Configurer votre nouveau projet**, définissez les champs suivants.</span><span class="sxs-lookup"><span data-stu-id="cbb82-190">In the **Configure your new project** dialog, set the following fields.</span></span>
    - <span data-ttu-id="cbb82-191">Donnez un **Nom de projet** à **CellAnalyzerSharedLibrary**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-191">Set the **Project name** to **CellAnalyzerSharedLibrary**.</span></span>
    - <span data-ttu-id="cbb82-192">Gardez l'**Emplacement** à sa valeur par défaut.</span><span class="sxs-lookup"><span data-stu-id="cbb82-192">Leave the **Location** at it's default value.</span></span>
    - <span data-ttu-id="cbb82-193">Configurez **Framework** sur **4.7.2**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-193">Set the **Framework** to **4.7.2**.</span></span>
5. <span data-ttu-id="cbb82-194">Sélectionnez **Créer**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-194">Choose **Create**.</span></span>
6. <span data-ttu-id="cbb82-195">Une fois le projet créé, renommez le fichier **Class1.cs** dans **CellOperations.cs**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-195">After the project is created, rename the **Class1.cs** file to **CellOperations.cs**.</span></span> <span data-ttu-id="cbb82-196">Une invite apparaît pour renommer la classe.</span><span class="sxs-lookup"><span data-stu-id="cbb82-196">A prompt to rename the class appears.</span></span> <span data-ttu-id="cbb82-197">Renommez le nom de classe pour qu’il corresponde au nom du fichier.</span><span class="sxs-lookup"><span data-stu-id="cbb82-197">Rename the class name so that it matches the file name.</span></span>
7. <span data-ttu-id="cbb82-198">Ajoutez le code suivant à la classe `CellOperations` pour créer une méthode nommée `GetUnicodeFromText`.</span><span class="sxs-lookup"><span data-stu-id="cbb82-198">Add the following code to the `CellOperations` class to create a method named `GetUnicodeFromText`.</span></span>

```csharp
public class CellOperations
{
    static public string GetUnicodeFromText(string value)
    {
        string result = "";
        foreach (char c in value)
        {
            int unicode = c;

            result += $"{c}: {unicode}\r\n";
        }
        return result;
    }
}
```

### <a name="use-the-shared-class-library-in-the-vsto-add-in"></a><span data-ttu-id="cbb82-199">Utiliser la bibliothèque de classes partagées dans le complément VSTO</span><span class="sxs-lookup"><span data-stu-id="cbb82-199">Use the shared class library in the VSTO Add-in</span></span>

<span data-ttu-id="cbb82-200">Vous devez maintenant mettre à jour le complément VSTO pour utiliser la bibliothèque de classes.</span><span class="sxs-lookup"><span data-stu-id="cbb82-200">Now you need to update the VSTO Add-in to use the class library.</span></span> <span data-ttu-id="cbb82-201">Il est important que les compléments VSTO et Office utilisent la même bibliothèque de classes partagées pour permettre de réaliser au même endroit les résolutions de bogues et les fonctionnalités.</span><span class="sxs-lookup"><span data-stu-id="cbb82-201">This is important that both the VSTO Add-in and Office Add-in use the same shared class library so that future bug fixes or features are made in one location.</span></span>

1. <span data-ttu-id="cbb82-202">Dans l’**Explorateur de solutions**, cliquez à l'aide du bouton droit sur le projet **Analyseur de cellules**, puis choisissez **Ajouter une référence**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-202">In **Solution Explorer** right-click the **Cell-Analyzer** project, and choose **Add Reference**.</span></span>
2. <span data-ttu-id="cbb82-203">Sélectionnez **CellAnalyzerSharedLibrary**, puis choisissez **OK**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-203">Select **CellAnalyzerSharedLibrary**, and choose **OK**.</span></span>
3. <span data-ttu-id="cbb82-204">Dans l'**Explorateur de solutions**, développez l'**Analyseur de cellules** du projet, cliquez avec le bouton droit sur le fichier **CellAnalyzerPane.cs**, puis sélectionnez **Afficher le code**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-204">In **Solution Explorer** expand the **Cell-Analyzer** project, right-click the **CellAnalyzerPane.cs** file, and choose **View Code**.</span></span>
4. <span data-ttu-id="cbb82-205">Dans la méthode `btnUnicode_Click`, supprimez les lignes de code suivantes.</span><span class="sxs-lookup"><span data-stu-id="cbb82-205">In the `btnUnicode_Click` method, delete the following lines of code.</span></span>

    ```csharp
    //Convert to Unicode listing
    string result = "";
    foreach (char c in cellValue)
    {
      int unicode = c;
      result += $"{c}: {unicode}\r\n";
    }
    ```

5. <span data-ttu-id="cbb82-206">Mettez à jour la ligne de code sous le commentaire à lire `//Output the result` comme suit :</span><span class="sxs-lookup"><span data-stu-id="cbb82-206">Update the line of code under the `//Output the result` comment to read as follows:</span></span>

    ```csharp
    //Output the result
    txtResult.Text = CellAnalyzerSharedLibrary.CellOperations.GetUnicodeFromText(cellValue);
    ```

6. <span data-ttu-id="cbb82-207">Dans le menu **Déboguer**, choisissez **Démarrer le débogage**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-207">On the **Debug** menu, choose **Start Debugging**.</span></span> <span data-ttu-id="cbb82-208">Le volet Office personnalisé doit fonctionner comme attendu.</span><span class="sxs-lookup"><span data-stu-id="cbb82-208">The custom task pane should work as expected.</span></span> <span data-ttu-id="cbb82-209">Entrez du texte dans une cellule, puis vérifiez que vous pouvez le convertir en liste Unicode avec le complément.</span><span class="sxs-lookup"><span data-stu-id="cbb82-209">Enter some text in a cell, and then test that you can convert it to a Unicode list with the add-in.</span></span>

## <a name="create-a-rest-api-wrapper"></a><span data-ttu-id="cbb82-210">Créer un wrapper API REST</span><span class="sxs-lookup"><span data-stu-id="cbb82-210">Create a REST API wrapper</span></span>

<span data-ttu-id="cbb82-211">Le complément VSTO peut utiliser directement la bibliothèque de classes partagée car tous deux sont des projets .NET.</span><span class="sxs-lookup"><span data-stu-id="cbb82-211">The VSTO Add-in can use the shared class library directly since they are both .NET projects.</span></span> <span data-ttu-id="cbb82-212">Le complément Office ne pourra toutefois pas utiliser .NET car il utilise JavaScript.</span><span class="sxs-lookup"><span data-stu-id="cbb82-212">However the Office Add-in won't be able to use .NET since it uses JavaScript.</span></span> <span data-ttu-id="cbb82-213">Vous devez ensuite créer un wrapper API REST.</span><span class="sxs-lookup"><span data-stu-id="cbb82-213">Next you will need to create a REST API wrapper.</span></span> <span data-ttu-id="cbb82-214">Le complément Office peut ainsi appeler une API REST, qui transmet ensuite l’appel vers la bibliothèque de classes partagée.</span><span class="sxs-lookup"><span data-stu-id="cbb82-214">This enables the Office Add-in to call a REST API, which then passes the call along to the shared class library.</span></span>

1. <span data-ttu-id="cbb82-215">Dans l’**Explorateur de solutions**, cliquez à l'aide du bouton droit sur le projet **Analyseur de cellules**, puis choisissez **Ajouter un nouveau projet**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-215">In **Solution Explorer**, right-click the **Cell-Analyzer** project, and choose **Add > New Project**.</span></span>
2. <span data-ttu-id="cbb82-216">Dans la **boîte de dialogue Ajouter un nouveau projet**, choisissez **Application web ASP.NET Core**, puis sélectionnez **Suivant**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-216">In the **Add a new project dialog**, choose **ASP.NET Core Web Application**, and choose **Next**.</span></span>
3. <span data-ttu-id="cbb82-217">Dans la boîte de dialogue **Configurer votre nouveau projet**, définissez les champs suivants :</span><span class="sxs-lookup"><span data-stu-id="cbb82-217">In the **Configure your new project** dialog, set the following fields:</span></span>
    - <span data-ttu-id="cbb82-218">Donnez un **Nom de projet** à **CellAnalyzerRESTAPI**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-218">Set the **Project name** to **CellAnalyzerRESTAPI**.</span></span>
    - <span data-ttu-id="cbb82-219">Dans le champ **Emplacement**, conserver la valeur par défaut.</span><span class="sxs-lookup"><span data-stu-id="cbb82-219">In the **Location** field, leave the default value.</span></span>
4. <span data-ttu-id="cbb82-220">Sélectionnez **Créer**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-220">Choose **Create**.</span></span>
5. <span data-ttu-id="cbb82-221">Dans la boîte de dialogue **Créer une application web ASP.NET Core**, sélectionnez **ASP.NET Core 3.1** pour la version, puis sélectionnez l'**API** dans la liste ds projets.</span><span class="sxs-lookup"><span data-stu-id="cbb82-221">In the **Create a new ASP.NET Core web application** dialog, select **ASP.NET Core 3.1** for the version, and select **API** in the list of projects.</span></span>
6. <span data-ttu-id="cbb82-222">Conservez les valeurs par défaut dans tous les autres champs, puis sélectionnez le bouton **Créer**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-222">Leave all other fields at default values and choose the **Create** button.</span></span>
7. <span data-ttu-id="cbb82-223">Une fois le projet créé, développez le projet **CellAnalyzerRESTAPI** dans l'**Explorateur de solutions**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-223">After the project is created, expand the **CellAnalyzerRESTAPI** project in **Solution Explorer**.</span></span>
8. <span data-ttu-id="cbb82-224">Cliquez avec le bouton droit sur **Dépendances**, puis sélectionnez **Ajouter une référence**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-224">Right-click **Dependencies**, and choose **Add Reference**.</span></span>
9. <span data-ttu-id="cbb82-225">Sélectionnez **CellAnalyzerSharedLibrary**, puis choisissez **OK**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-225">Select **CellAnalyzerSharedLibrary**, and choose **OK**.</span></span>
10. <span data-ttu-id="cbb82-226">Cliquez avec le bouton droit sur le dossier **Contrôleurs**, puis choisissez **Ajouter > Contrôleur**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-226">Right-click the **Controllers** folder, and choose **Add > Controller**.</span></span>
11. <span data-ttu-id="cbb82-227">Dans la boîte de dialogue **Ajouter un nouvel élément structuré**, sélectionnez **Contrôleur d'API – Vide**, puis **Ajouter**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-227">In the **Add New Scaffolded Item** dialog, choose **API Controller - Empty** and then **Add**.</span></span>
12. <span data-ttu-id="cbb82-228">Dans la boîte de dialogue **Ajouter un contrôleur d'API vide**, nommez le contrôleur **AnalyzeUnicodeController**, puis sélectionnez **Ajouter**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-228">In the **Add Empty API Controller** dialog, name the controller **AnalyzeUnicodeController**, and then choose **Add**.</span></span>
13. <span data-ttu-id="cbb82-229">Ouvrez le fichier **AnalyzeUnicodeController.cs** et ajoutez le code suivant en tant que méthode à la classe `AnalyzeUnicodeController`.</span><span class="sxs-lookup"><span data-stu-id="cbb82-229">Open the **AnalyzeUnicodeController.cs** file and add the following code as a method to the `AnalyzeUnicodeController` class.</span></span>

    ```csharp
    [HttpGet]
    public ActionResult<string> AnalyzeUnicode(string value)
    {
      if (value == null)
      {
        return BadRequest();
      }
      return CellAnalyzerSharedLibrary.CellOperations.GetUnicodeFromText(value);
    }
    ```

14. <span data-ttu-id="cbb82-230">Cliquez avec le bouton droit sur le projet **CellAnalyzerRESTAPI**, puis choisissez **Définir comme projet de démarrage**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-230">Right-click the **CellAnalyzerRESTAPI** project, and choose **Set as Startup Project**.</span></span>
15. <span data-ttu-id="cbb82-231">Dans le menu **Déboguer**, choisissez **Démarrer le débogage**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-231">On the **Debug** menu, choose **Start Debugging**.</span></span>
16. <span data-ttu-id="cbb82-232">Un navigateur s’ouvre.</span><span class="sxs-lookup"><span data-stu-id="cbb82-232">A browser will launch.</span></span> <span data-ttu-id="cbb82-233">Entrez l’URL suivante pour vérifier que l’API REST fonctionne : `https://localhost:<ssl port number>/api/analyzeunicode?value=test`.</span><span class="sxs-lookup"><span data-stu-id="cbb82-233">Enter the following URL to test that the REST API is working: `https://localhost:<ssl port number>/api/analyzeunicode?value=test`.</span></span> <span data-ttu-id="cbb82-234">Vous pouvez réutiliser le numéro de port à partir de l’URL dans le navigateur lancé par Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="cbb82-234">You can reuse the port number from the URL in the browser that Visual Studio launched.</span></span> <span data-ttu-id="cbb82-235">Vous devriez voir une chaîne renvoyée avec des valeurs Unicode pour chaque caractère.</span><span class="sxs-lookup"><span data-stu-id="cbb82-235">You should see a string returned with Unicode values for each character.</span></span>

## <a name="create-the-office-add-in"></a><span data-ttu-id="cbb82-236">Créer le complément Office</span><span class="sxs-lookup"><span data-stu-id="cbb82-236">Create the Office Add-in</span></span>

<span data-ttu-id="cbb82-237">Lorsque vous créez le complément Office, celui-ci appelle l'API REST.</span><span class="sxs-lookup"><span data-stu-id="cbb82-237">When you create the Office Add-in, it will make a call to the REST API.</span></span> <span data-ttu-id="cbb82-238">Mais vous devez tout d'abord obtenir le numéro de port du serveur API REST et de l’enregistrer pour plus tard.</span><span class="sxs-lookup"><span data-stu-id="cbb82-238">But first, you need to get the port number of the REST API server and save it for later.</span></span>

### <a name="save-the-ssl-port-number"></a><span data-ttu-id="cbb82-239">Enregistrer le numéro de port SSL</span><span class="sxs-lookup"><span data-stu-id="cbb82-239">Save the SSL port number</span></span>

1. <span data-ttu-id="cbb82-240">Si ce n'est pas encore fait, démarrez Visual Studio 2019 et ouvrez la solution **\start\Cell-Analyzer.sln**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-240">If you haven't already, start Visual Studio 2019, and open the **\start\Cell-Analyzer.sln** solution.</span></span>
2. <span data-ttu-id="cbb82-241">Dans le projet **CellAnalyzerRESTAPI**, développez les **Propriétés** et ouvrez le fichier **launchSettings.json**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-241">In the **CellAnalyzerRESTAPI** project, expand **Properties**, and open the **launchSettings.json** file.</span></span>
3. <span data-ttu-id="cbb82-242">Recherchez la ligne de code contenant la valeur de **sslPort**, copiez le numéro de port et enregistrez-le quelque part.</span><span class="sxs-lookup"><span data-stu-id="cbb82-242">Find the line of code with the **sslPort** value, copy the port number, and save it somewhere.</span></span>

### <a name="add-the-office-add-in-project"></a><span data-ttu-id="cbb82-243">Ajouter le projet de complément Office</span><span class="sxs-lookup"><span data-stu-id="cbb82-243">Add the Office Add-in project</span></span>

<span data-ttu-id="cbb82-244">Pour simplifier les choses, conservez tous les codes dans une seule solution.</span><span class="sxs-lookup"><span data-stu-id="cbb82-244">To keep things simple, keep all the code in one solution.</span></span> <span data-ttu-id="cbb82-245">Ajoutez le projet de complément Office à la solution Visual Studio existante.</span><span class="sxs-lookup"><span data-stu-id="cbb82-245">Add the Office Add-in project to the existing Visual Studio solution.</span></span> <span data-ttu-id="cbb82-246">Toutefois, si vous avez l’habitude d’utiliser le [Générateur Yeoman pour compléments Office](https://github.com/OfficeDev/generator-office) et Visual Studio Code, vous pouvez également exécuter `yo office` pour générer le projet.</span><span class="sxs-lookup"><span data-stu-id="cbb82-246">However, if you are familiar with the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) and Visual Studio Code you can also run `yo office` to build the project.</span></span> <span data-ttu-id="cbb82-247">Les étapes sont très semblables.</span><span class="sxs-lookup"><span data-stu-id="cbb82-247">The steps are very similar.</span></span>

1. <span data-ttu-id="cbb82-248">Dans l’**Explorateur de solutions**, cliquez à l'aide du bouton droit sur la solution **Analyseur de cellules**, puis choisissez **Ajouter > Nouveau projet**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-248">In **Solution Explorer**, right-click the **Cell-Analyzer** solution, and choose **Add > New Project**.</span></span>
2. <span data-ttu-id="cbb82-249">Dans la **Boîte de dialogue Ajouter un nouveau projet**, choisissez **Complément web Excel**, puis sélectionnez **Suivant**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-249">In the **Add a new project dialog**, choose **Excel Web Add-in**, and choose **Next**.</span></span>
3. <span data-ttu-id="cbb82-250">Dans la boîte de dialogue **Configurer votre nouveau projet**, définissez les champs suivants :</span><span class="sxs-lookup"><span data-stu-id="cbb82-250">In the **Configure your new project** dialog, set the following fields:</span></span>
    - <span data-ttu-id="cbb82-251">Donnez un **Nom de projet** à **CellAnalyzerOfficeAddin**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-251">Set the **Project name** to **CellAnalyzerOfficeAddin**.</span></span>
    - <span data-ttu-id="cbb82-252">Gardez l'**Emplacement** à sa valeur par défaut.</span><span class="sxs-lookup"><span data-stu-id="cbb82-252">Leave the **Location** at it's default value.</span></span>
    - <span data-ttu-id="cbb82-253">Configurez **Framework** sur **4.7.2** ou une version ultérieure.</span><span class="sxs-lookup"><span data-stu-id="cbb82-253">Set the **Framework** to **4.7.2** or later.</span></span>
4. <span data-ttu-id="cbb82-254">Sélectionnez **Créer**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-254">Choose **Create**.</span></span>
5. <span data-ttu-id="cbb82-255">Dans la boîte de dialogue **Choisir le type de complément**, sélectionnez **Ajouter e nouvelles fonctionnalités dans Excel**, puis choisissez **Terminer**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-255">In the **Choose the add-in type** dialog, select **Add new functionalities to Excel**, and choose **Finish**.</span></span>

<span data-ttu-id="cbb82-256">Deux projets sont créés :</span><span class="sxs-lookup"><span data-stu-id="cbb82-256">Two projects will be created:</span></span>

- <span data-ttu-id="cbb82-257">**CellAnalyzerOfficeAddin** : ce projet configure les fichiers XML du manifeste qui décrivent le complément pour qu’Office puisse le charger correctement.</span><span class="sxs-lookup"><span data-stu-id="cbb82-257">**CellAnalyzerOfficeAddin** - This project configures the manifest XML files that describes the add-in so Office can load it correctly.</span></span> <span data-ttu-id="cbb82-258">Il contient l’ID, le nom, la description et d’autres informations sur le complément.</span><span class="sxs-lookup"><span data-stu-id="cbb82-258">It contains the ID, name, description, and other information about the add-in.</span></span>
- <span data-ttu-id="cbb82-259">**CellAnalyzerOfficeAddinWeb** : ce projet contient des ressources Web pour votre complément (par exemple, HTML, CSS et des scripts).</span><span class="sxs-lookup"><span data-stu-id="cbb82-259">**CellAnalyzerOfficeAddinWeb** - This project contains web resources for your add-in, such as HTML, CSS, and scripts.</span></span> <span data-ttu-id="cbb82-260">Il configure également une instance IIS Express pour héberger votre complément en tant qu’application Web.</span><span class="sxs-lookup"><span data-stu-id="cbb82-260">It also configures an IIS Express instance to host your add-in as a web application.</span></span>

### <a name="add-ui-and-functionality-to-the-office-add-in"></a><span data-ttu-id="cbb82-261">Ajouter des interfaces utilisateur et des fonctionnalités au complément Office</span><span class="sxs-lookup"><span data-stu-id="cbb82-261">Add UI and functionality to the Office Add-in</span></span>

1. <span data-ttu-id="cbb82-262">Dans l'**Explorateur de solutions**, développez le projet **CellAnalyzerOfficeAddinWeb**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-262">In **Solution Explorer**, expand the **CellAnalyzerOfficeAddinWeb** project.</span></span>
2. <span data-ttu-id="cbb82-263">Ouvrez le fichier **Home.html** et remplacez le contenu `<body>` par l'HTML suivant.</span><span class="sxs-lookup"><span data-stu-id="cbb82-263">Open the **Home.html** file, and replace the `<body>` contents with the following HTML.</span></span>

    ```html
    <button id="btnShowUnicode" onclick="showUnicode()">Show Unicode</button>
    <p>Result:</p>
    <div id="txtResult"></div>
    ```

3. <span data-ttu-id="cbb82-264">Ouvrez ce fichier **Home.js** et remplacez l’intégralité de son contenu par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="cbb82-264">Open the **Home.js** file, and replace the entire contents with the following code.</span></span>

    ```js
    (function () {
      "use strict";
      // The initialize function must be run each time a new page is loaded.
      Office.initialize = function (reason) {
        $(document).ready(function () {
        });
      };
    })();

    function showUnicode() {
      Excel.run(function (ctx) {
        const range = ctx.workbook.getSelectedRange();
        range.load("values");
        return ctx.sync(range).then(function (range) {
          const url = "https://localhost:<ssl port number>/api/analyzeunicode?value=" + range.values[0][0];
          $.ajax({
            type: "GET",
            url: url,
            success: function (data) {
              let htmlData = data.replace(/\r\n/g, '<br>');
              $("#txtResult").html(htmlData);
            },
            error: function (data) {
                $("#txtResult").html("error occurred in ajax call.");
            }
          });
        });
      });
    }
    ```

4. <span data-ttu-id="cbb82-265">Dans le code précédent, entrez le numéro de **sslPort** que vous avez enregistré précédemment à partir du fichier **launchSettings.json**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-265">In the previous code, enter the **sslPort** number you saved previously from the **launchSettings.json** file.</span></span>

<span data-ttu-id="cbb82-266">Dans le code précédent, la chaîne renvoyée est traitée pour remplacer les sauts de ligne avec retour chariot par des balises HTML `<br>`.</span><span class="sxs-lookup"><span data-stu-id="cbb82-266">In the previous code the returned string will be processed to replace carriage return line feeds with `<br>` HTML tags.</span></span> <span data-ttu-id="cbb82-267">Vous pouvez parfois être confronté(e) à des situations dans lesquelles une valeur de retour fonctionnant parfaitement pour .NET dans le complément VSTO doit être ajustée sur le côté du complément Office pour fonctionner comme attendu.</span><span class="sxs-lookup"><span data-stu-id="cbb82-267">You may occasionally run into situations where a return value that works perfectly fine for .NET in the VSTO Add-in will need to be adjusted on the Office Add-in side to work as expected.</span></span> <span data-ttu-id="cbb82-268">Dans ce cas, l’API REST et la bibliothèque de classes partagées s'intéressent uniquement au retour de chaîne.</span><span class="sxs-lookup"><span data-stu-id="cbb82-268">In this case the REST API and shared class library are only concerned with returning the string.</span></span> <span data-ttu-id="cbb82-269">La méthode `showUnicode()` est chargée de la mise en forme correcte des valeurs de retour pour la présentation.</span><span class="sxs-lookup"><span data-stu-id="cbb82-269">The `showUnicode()` method is responsible for formatting return values correctly for presentation.</span></span>

### <a name="allow-cors-from-the-office-add-in"></a><span data-ttu-id="cbb82-270">Autoriser CORS à partir d'un complément Office</span><span class="sxs-lookup"><span data-stu-id="cbb82-270">Allow CORS from the Office Add-in</span></span>

<span data-ttu-id="cbb82-271">La bibliothèque Office.js nécessite CORS pour les appels sortants, tels que ceux effectués à partir de l’appel `ajax` vers le serveur API REST.</span><span class="sxs-lookup"><span data-stu-id="cbb82-271">The Office.js library requires CORS on outgoing calls, such as the one made from the `ajax` call to the REST API server.</span></span> <span data-ttu-id="cbb82-272">Pour autoriser des appels du complément Office vers l’API REST, procédez comme suit.</span><span class="sxs-lookup"><span data-stu-id="cbb82-272">Use the following steps to allow calls from the Office Add-in to the REST API.</span></span>

1. <span data-ttu-id="cbb82-273">Dans l'**Explorateur de solutions**, sélectionnez le projet **CellAnalyzerOfficeAddinWeb**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-273">In **Solution Explorer**, select the **CellAnalyzerOfficeAddinWeb** project.</span></span>
2. <span data-ttu-id="cbb82-274">Dans le menu **Afficher**, Choisissez **Fenêtre des Propriétés** (si la fenêtre ne s'affiche pas).</span><span class="sxs-lookup"><span data-stu-id="cbb82-274">From the **View** menu, choose **Properties Window** (if the window is not already displayed).</span></span>
3. <span data-ttu-id="cbb82-275">Dans la fenêtre des propriétés, copiez et enregistrez la valeur de l'**URL SSL**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-275">In the properties window, copy the value of the **SSL URL**, and save it somewhere.</span></span> <span data-ttu-id="cbb82-276">Il s’agit de l’URL que vous devez autoriser dans CORS.</span><span class="sxs-lookup"><span data-stu-id="cbb82-276">This is the URL that you need to allow through CORS.</span></span>
4. <span data-ttu-id="cbb82-277">Dans le projet **CellAnalyzerRESTAPI**, ouvrez le fichier **Startup.cs**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-277">In the **CellAnalyzerRESTAPI** project, open the **Startup.cs** file.</span></span>
5. <span data-ttu-id="cbb82-278">Ajoutez le code suivant en haut de la méthode `ConfigureServices`.</span><span class="sxs-lookup"><span data-stu-id="cbb82-278">Add the following code to the top of the `ConfigureServices` method.</span></span> <span data-ttu-id="cbb82-279">Assurez-vous de remplacer l’URL SSL que vous avez copiée précédemment pour l’appel `builder.WithOrigins`.</span><span class="sxs-lookup"><span data-stu-id="cbb82-279">Be sure to substitute the URL SSL you copied previously for the `builder.WithOrigins` call.</span></span>

    ```csharp
    services.AddCors(options =>
    {
      options.AddPolicy(MyAllowSpecificOrigins,
      builder =>
      {
        builder.WithOrigins("<your URL SSL>")
        .AllowAnyMethod()
        .AllowAnyHeader();
      });
    });
    ```

    > [!NOTE]
    > <span data-ttu-id="cbb82-280">Enlevez le `/` qui se trouve à la fin de l’URL lorsque vous l’utilisez dans la méthode `builder.WithOrigins`Builder.WithOrigins.tr.</span><span class="sxs-lookup"><span data-stu-id="cbb82-280">Leave the trailing `/` from the end of the URL when you use it in the `builder.WithOrigins` method.</span></span> <span data-ttu-id="cbb82-281">Par exemple, il doit ressembler à `https://localhost:44000`.</span><span class="sxs-lookup"><span data-stu-id="cbb82-281">For example, it should appear similar to `https://localhost:44000`.</span></span> <span data-ttu-id="cbb82-282">Dans le cas contraire, une erreur CORS se produira lors de l’exécution.</span><span class="sxs-lookup"><span data-stu-id="cbb82-282">Otherwise you will get a CORS error at runtime.</span></span>

6. <span data-ttu-id="cbb82-283">Ajoutez le champs suivant à la classe `Startup` :</span><span class="sxs-lookup"><span data-stu-id="cbb82-283">Add the following field to the `Startup` class:</span></span>

    ```csharp
    readonly string MyAllowSpecificOrigins = "_myAllowSpecificOrigins";
    ```

7. <span data-ttu-id="cbb82-284">Ajoutez le code suivant à la méthode `configure` juste avant la ligne de code pour `app.UseEndpoints`.</span><span class="sxs-lookup"><span data-stu-id="cbb82-284">Add the following code to the `configure` method just before the line of code for `app.UseEndpoints`.</span></span>

    ```csharp
    app.UseCors(MyAllowSpecificOrigins);
    ```

<span data-ttu-id="cbb82-285">Lorsque vous avez terminé, votre classe `Startup` doit ressembler au code suivant (votre URL localhost peut être différente) :</span><span class="sxs-lookup"><span data-stu-id="cbb82-285">When done, your `Startup` class should look similar to the following code (your localhost URL may be different):</span></span>

```csharp
public class Startup
{
  public Startup(IConfiguration configuration)
    {
      Configuration = configuration;
    }

    readonly string MyAllowSpecificOrigins = "_myAllowSpecificOrigins";

    public IConfiguration Configuration { get; }

    // NOTE: The following code configures CORS for the localhost:44397 port.
    // This is for development purposes. In production code you should update this to 
    // use the appropriate allowed domains.
    public void ConfigureServices(IServiceCollection services)
    {
        services.AddCors(options =>
        {
            options.AddPolicy(MyAllowSpecificOrigins,
            builder =>
            {
                builder.WithOrigins("https://localhost:44397")
                .AllowAnyMethod()
                .AllowAnyHeader();
            });
        });
        services.AddControllers();
    }

    // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
    public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
    {
        if (env.IsDevelopment())
        {
            app.UseDeveloperExceptionPage();
        }

        app.UseHttpsRedirection();

        app.UseRouting();

        app.UseAuthorization();

        app.UseCors(MyAllowSpecificOrigins);

        app.UseEndpoints(endpoints =>
        {
            endpoints.MapControllers();
        });
    }
}
```

### <a name="run-the-add-in"></a><span data-ttu-id="cbb82-286">Exécuter du complément</span><span class="sxs-lookup"><span data-stu-id="cbb82-286">Run the add-in</span></span>

1. <span data-ttu-id="cbb82-287">Dans l’**Explorateur de solutions**, cliquez à l'aide du nœud supérieur sur la **Solution de l'analyseur de cellules**, puis choisissez **Configurer les projets de départ**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-287">In **Solution Explorer**, right-click the top node **Solution 'Cell-Analyzer'**, and choose **Set Startup Projects**.</span></span>
2. <span data-ttu-id="cbb82-288">Dans la boîte de dialogue des **Pages de propriété de la solution de l'analyseur de cellules**, sélectionnez **Plusieurs projets de départ**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-288">In the **Solution 'Cell-Analyzer' Property Pages** dialog, select **Multiple startup projects**.</span></span>
3. <span data-ttu-id="cbb82-289">Définissez la propriété **action** au **Départ** pour chacun des projets suivants.</span><span class="sxs-lookup"><span data-stu-id="cbb82-289">Set the **Action** property to **Start** for each of the following projects.</span></span>

    - <span data-ttu-id="cbb82-290">CellAnalyzerRESTAPI</span><span class="sxs-lookup"><span data-stu-id="cbb82-290">CellAnalyzerRESTAPI</span></span>
    - <span data-ttu-id="cbb82-291">CellAnalyzerOfficeAddin</span><span class="sxs-lookup"><span data-stu-id="cbb82-291">CellAnalyzerOfficeAddin</span></span>
    - <span data-ttu-id="cbb82-292">CellAnalyzerOfficeAddinWeb</span><span class="sxs-lookup"><span data-stu-id="cbb82-292">CellAnalyzerOfficeAddinWeb</span></span>

4. <span data-ttu-id="cbb82-293">Sélectionnez **OK**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-293">Choose **OK**.</span></span>
5. <span data-ttu-id="cbb82-294">Dans le menu **Déboguer**, choisissez **Démarrer le débogage**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-294">From the **Debug** menu, choose **Start Debugging**.</span></span>

<span data-ttu-id="cbb82-295">Excel exécute et charge une version test du complément Office.</span><span class="sxs-lookup"><span data-stu-id="cbb82-295">Excel will run and sideload the Office Add-in.</span></span> <span data-ttu-id="cbb82-296">Vous pouvez vérifier que le service API REST localhost fonctionne correctement en entrant une valeur de texte dans une cellule, puis en sélectionnant le bouton **Afficher l'Unicode** dans le complément Office.</span><span class="sxs-lookup"><span data-stu-id="cbb82-296">You can test that the localhost REST API service is working correctly by entering a text value into a cell, and choosing the **Show Unicode** button in the Office Add-in.</span></span> <span data-ttu-id="cbb82-297">Il doit appeler l’API REST et afficher les valeurs Unicode pour les caractères de texte.</span><span class="sxs-lookup"><span data-stu-id="cbb82-297">It should call the REST API and display the unicode values for the text characters.</span></span>

## <a name="publish-to-an-azure-app-service"></a><span data-ttu-id="cbb82-298">Publier vers Azure App Service</span><span class="sxs-lookup"><span data-stu-id="cbb82-298">Publish to an Azure App Service</span></span>

<span data-ttu-id="cbb82-299">Vous voulez enfin publier le projet API REST sur le cloud.</span><span class="sxs-lookup"><span data-stu-id="cbb82-299">You eventually want to publish the REST API project to the cloud.</span></span> <span data-ttu-id="cbb82-300">Dans les étapes suivantes, vous allez découvrir comment publier le projet **CellAnalyzerRESTAPI** dans Microsoft Azure App Service.</span><span class="sxs-lookup"><span data-stu-id="cbb82-300">In the following steps you'll see how to publish the **CellAnalyzerRESTAPI** project to a Microsoft Azure App Service.</span></span> <span data-ttu-id="cbb82-301">Pour plus d’informations sur l’obtention d’un compte Azure, voir les [Conditions préalables](#prerequisites).</span><span class="sxs-lookup"><span data-stu-id="cbb82-301">See [Prerequisites](#prerequisites) for information on how to get an Azure account.</span></span>

1. <span data-ttu-id="cbb82-302">Dans l’**Explorateur de solutions**, cliquez à l'aide du bouton droit sur le projet **CellAnalyzerRESTAPI**, puis choisissez **Publier**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-302">In **Solution Explorer**, right-click the **CellAnalyzerRESTAPI** project, and choose **Publish**.</span></span>
2. <span data-ttu-id="cbb82-303">Dans la boîte de dialogue **Sélectionner une cible de publication**, sélectionnez **Créer nouveau**, puis choisissez **Créer un profil**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-303">In the **Pick a publish target** dialog, select **Create New**, and choose **Create Profile**.</span></span>
3. <span data-ttu-id="cbb82-304">Dans la boîte de dialogue **App Service**, sélectionnez le compte correct, s’il n’est pas encore choisi.</span><span class="sxs-lookup"><span data-stu-id="cbb82-304">In the **App Service** dialog, select the correct account, if it is not already selected.</span></span>
4. <span data-ttu-id="cbb82-305">Les valeurs par défaut sont utilisées dans les champs de la boîte de dialogue **App Service** de votre compte.</span><span class="sxs-lookup"><span data-stu-id="cbb82-305">The fields for the **App Service** dialog will be set to defaults for your account.</span></span> <span data-ttu-id="cbb82-306">Les valeurs par défaut fonctionnent correctement en général, mais vous pouvez les modifier si vous préférez définir d’autres paramètres.</span><span class="sxs-lookup"><span data-stu-id="cbb82-306">Generally the defaults work fine, but you can change them if you prefer different settings.</span></span>
5. <span data-ttu-id="cbb82-307">Dans la boîte de dialogue **App Service**, sélectionnez **Créer**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-307">In the **App Service** dialog, choose **Create**.</span></span>
6. <span data-ttu-id="cbb82-308">Le nouveau profil s’affiche dans une page **Publier**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-308">The new profile will be displayed in a **Publish** page.</span></span> <span data-ttu-id="cbb82-309">Sélectionnez **Publier** pour créer et déployer le code vers App Service.</span><span class="sxs-lookup"><span data-stu-id="cbb82-309">Choose **Publish** to build and deploy the code to the App Service.</span></span>

<span data-ttu-id="cbb82-310">Vous pouvez maintenant tester le service.</span><span class="sxs-lookup"><span data-stu-id="cbb82-310">You can now test the service.</span></span> <span data-ttu-id="cbb82-311">Ouvrez un navigateur et entrez une URL qui accède directement au nouveau service.</span><span class="sxs-lookup"><span data-stu-id="cbb82-311">Open a browser and enter a URL that goes directly to the new service.</span></span> <span data-ttu-id="cbb82-312">Par exemple, utilisez `https://<myappservice>.azurewebsites.net/api/analyzeunicode?value=test` où *myappservice* est le seul nom que vous avez créé pour le nouvel App Service.</span><span class="sxs-lookup"><span data-stu-id="cbb82-312">For example, use `https://<myappservice>.azurewebsites.net/api/analyzeunicode?value=test` where *myappservice* is the unique name you created for the new App Service.</span></span>

### <a name="use-the-azure-app-service-from-the-office-add-in"></a><span data-ttu-id="cbb82-313">Utiliser Azure App Service à partir du complément Office</span><span class="sxs-lookup"><span data-stu-id="cbb82-313">Use the Azure App Service from the Office Add-in</span></span>

<span data-ttu-id="cbb82-314">La dernière étape consiste à mettre à jour le code dans le complément Office pour utiliser Azure App Service au lieu de localhost.</span><span class="sxs-lookup"><span data-stu-id="cbb82-314">The final step is to update the code in the Office Add-in to use the Azure App Service instead of localhost.</span></span>

1. <span data-ttu-id="cbb82-315">Dans l'**Explorateur de solutions**, développez le projet **CellAnalyzerOfficeAddinWeb** et ouvrez le fichier **Home.js**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-315">In **Solution Explorer**, expand the **CellAnalyzerOfficeAddinWeb** project, and open the **Home.js** file.</span></span>
1. <span data-ttu-id="cbb82-316">Modifiez la constante `url` afin d’utiliser l’URL d'Azure App Service, comme illustré dans la ligne de code suivante.</span><span class="sxs-lookup"><span data-stu-id="cbb82-316">Change the `url` constant to use the URL for your Azure App Service as shown in the following line of code.</span></span> <span data-ttu-id="cbb82-317">Remplacez `<myappservice>` par le nom unique que vous avez créé pour le nouvel App Service.</span><span class="sxs-lookup"><span data-stu-id="cbb82-317">Replace `<myappservice>` with the unique name you created for the new App Service.</span></span>

    ```JavaScript
    const url = "https://<myappservice>.azurewebsites.net/api/analyzeunicode?value=" + range.values[0][0];
    ```

1. <span data-ttu-id="cbb82-318">Dans l’**Explorateur de solutions**, cliquez à l'aide du nœud supérieur sur la **Solution de l'analyseur de cellules**, puis choisissez **Configurer les projets de départ**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-318">In **Solution Explorer**, right-click the top node **Solution 'Cell-Analyzer'**, and choose **Set Startup Projects**.</span></span>
1. <span data-ttu-id="cbb82-319">Dans la boîte de dialogue des **Pages de propriété de la solution de l'analyseur de cellules**, sélectionnez **Plusieurs projets de départ**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-319">In the **Solution 'Cell-Analyzer' Property Pages** dialog, select **Multiple startup projects**.</span></span>
1. <span data-ttu-id="cbb82-320">Activez l’action **Démarrer** pour chacun des projets suivants :</span><span class="sxs-lookup"><span data-stu-id="cbb82-320">Enable the **Start** action for each of the following projects:</span></span>
    - <span data-ttu-id="cbb82-321">CellAnalyzerOfficeAddinWeb</span><span class="sxs-lookup"><span data-stu-id="cbb82-321">CellAnalyzerOfficeAddinWeb</span></span>
    - <span data-ttu-id="cbb82-322">CellAnalyzerOfficeAddin</span><span class="sxs-lookup"><span data-stu-id="cbb82-322">CellAnalyzerOfficeAddin</span></span>
1. <span data-ttu-id="cbb82-323">Sélectionnez **OK**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-323">Choose **OK**.</span></span>
1. <span data-ttu-id="cbb82-324">Dans le menu **Déboguer**, choisissez **Démarrer le débogage**.</span><span class="sxs-lookup"><span data-stu-id="cbb82-324">From the **Debug** menu, choose **Start Debugging**.</span></span>

<span data-ttu-id="cbb82-325">Excel exécute et charge une version test du complément Office.</span><span class="sxs-lookup"><span data-stu-id="cbb82-325">Excel will run and sideload the Office Add-in.</span></span> <span data-ttu-id="cbb82-326">Pour vérifier que App Service fonctionne correctement, entrez une valeur de texte dans une cellule, puis choisissez **Afficher l'Unicode** dans le complément Office.</span><span class="sxs-lookup"><span data-stu-id="cbb82-326">To test that the App Service is working correctly, enter a text value into a cell, and choose **Show Unicode** in the Office Add-in.</span></span> <span data-ttu-id="cbb82-327">Il doit appeler le service et afficher les valeurs Unicode pour les caractères de texte.</span><span class="sxs-lookup"><span data-stu-id="cbb82-327">It should call the service and display the unicode values for the text characters.</span></span>

## <a name="conclusion"></a><span data-ttu-id="cbb82-328">Conclusion</span><span class="sxs-lookup"><span data-stu-id="cbb82-328">Conclusion</span></span>

<span data-ttu-id="cbb82-329">Dans ce didacticiel, vous avez appris à créer un complément Office qui utilise un code partagé avec le complément VSTO d’origine.</span><span class="sxs-lookup"><span data-stu-id="cbb82-329">In this tutorial you learned how to create an Office Add-in that uses shared code with the original VSTO add-in.</span></span> <span data-ttu-id="cbb82-330">Vous avez appris à gérer le code VSTO pour Office sur Windows et un complément Office pour Office sur d’autres plateformes.</span><span class="sxs-lookup"><span data-stu-id="cbb82-330">You learned how to maintain both VSTO code for Office on Windows, and an Office Add-in for Office on other platforms.</span></span> <span data-ttu-id="cbb82-331">Vous avez refactorisé un code C# VSTO dans une bibliothèque partagée et vous l’avez déployé dans Azure App Service.</span><span class="sxs-lookup"><span data-stu-id="cbb82-331">You refactored VSTO C# code into a shared library and deployed it to an Azure App Service.</span></span> <span data-ttu-id="cbb82-332">Vous avez créé un complément Office qui utilise la bibliothèque partagée pour que vous n’ayez pas à réécrire le code dans JavaScript.</span><span class="sxs-lookup"><span data-stu-id="cbb82-332">You created an Office Add-in that uses the shared library so that you don't have to rewrite the code in JavaScript.</span></span>
