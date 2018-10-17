---
title: Débogage des compléments avec les outils de développement F12 sur Windows 10
description: ''
ms.date: 10/16/2018
ms.openlocfilehash: 3df245fcd651ec227e0a32d53da186ee332beb8f
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579841"
---
# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a><span data-ttu-id="edbe9-102">Débogage des compléments avec les outils de développement F12 sur Windows 10</span><span class="sxs-lookup"><span data-stu-id="edbe9-102">Debug add-ins using F12 developer tools on Windows 10</span></span>

<span data-ttu-id="edbe9-p101">Les outils de développement F12 inclus dans Windows 10 vous aident à déboguer, tester et accélérer vos pages web. Ils vous aident également à développer et déboguer les compléments Office si vous n’utilisez pas un IDE comme Visual Studio ou si vous devez examiner un problème pendant l’exécution de votre complément hors de l’IDE.  Dans cet article, vous découvrirez comment utiliser le débogueur des outils de développement F12 de Windows 10 pour tester votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="edbe9-p101">The F12 developer tools included in Windows 10 help you debug, test, and speed up your webpages. You can also use them to develop and debug Office Add-ins, if you are not using an IDE like Visual Studio, or if you need to investigate a problem while running your add-in outside the IDE. You can start the F12 developer tools after your add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="edbe9-106">Les instructions fournies dans cet article ne peuvent pas être utilisées pour déboguer un complément Outlook qui utilise des fonctions d’exécution.</span><span class="sxs-lookup"><span data-stu-id="edbe9-106">The instructions in this article cannot be used to debug an Outlook add-in that uses Execute Functions.</span></span> <span data-ttu-id="edbe9-107">Pour déboguer un complément Outlook qui utilise des fonctions d’exécution, nous vous recommandons de vous connecter à Visual Studio en mode script ou à un autre débogueur de script.</span><span class="sxs-lookup"><span data-stu-id="edbe9-107">To debug an Outlook add-in that uses Execute Functions, we recommend that you attach to Visual Studio in script mode or to some other script debugger.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="edbe9-108">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="edbe9-108">Prerequisites</span></span>

<span data-ttu-id="edbe9-109">Les logiciels suivants doivent être installés :</span><span class="sxs-lookup"><span data-stu-id="edbe9-109">You need the following software:</span></span>

- <span data-ttu-id="edbe9-110">Les outils de développement F12, inclus dans Windows 10.</span><span class="sxs-lookup"><span data-stu-id="edbe9-110">The F12 developer tools, which are included in Windows 10.</span></span> 
    
- <span data-ttu-id="edbe9-111">L’application cliente Office qui héberge votre complément.</span><span class="sxs-lookup"><span data-stu-id="edbe9-111">The Office client application that hosts your add-in.</span></span> 
    
- <span data-ttu-id="edbe9-112">Votre complément.</span><span class="sxs-lookup"><span data-stu-id="edbe9-112">Your add-in.</span></span> 

## <a name="using-the-debugger"></a><span data-ttu-id="edbe9-113">Utilisation du débogueur</span><span class="sxs-lookup"><span data-stu-id="edbe9-113">Using the Debugger</span></span>

<span data-ttu-id="edbe9-114">Vous pouvez utiliser le débogueur des outils de développement F12  de Windows 10 pour tester les compléments d’AppSource ou les compléments que vous avez ajoutés à partir d’autres emplacements.</span><span class="sxs-lookup"><span data-stu-id="edbe9-114">This article shows how you how to use the Debugger tool from the F12 developer tools in Windows 10 to test your Office Add-in. You can test add-ins from AppSource or add-ins that you have added from other locations. The F12 tools display in a separate window and do not use Visual Studio.</span></span> <span data-ttu-id="edbe9-115">Vous pouvez démarrer les outils de développement F12 après l’exécution de votre complément.</span><span class="sxs-lookup"><span data-stu-id="edbe9-115">You can start the F12 developer tools after your add-in is running.</span></span> <span data-ttu-id="edbe9-116">Les outils F12 s’ouvrent dans une fenêtre séparée et n’utilisent pas Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="edbe9-116">The F12 tools display in a separate window and do not use Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="edbe9-p104">Le débogueur fait partie des outils de développement F12 de Windows 10 et d’Internet Explorer. Il n’est pas inclus dans les versions antérieures de Windows.</span><span class="sxs-lookup"><span data-stu-id="edbe9-p104">The Debugger is part of the F12 developer tools in Windows 10 and Internet Explorer. Earlier versions of Windows do not include the Debugger.</span></span> 

<span data-ttu-id="edbe9-119">Cet exemple utilise Word et un complément gratuit d’AppSource.</span><span class="sxs-lookup"><span data-stu-id="edbe9-119">This example uses Word and a free add-in from AppSource.</span></span>

1. <span data-ttu-id="edbe9-120">Ouvrez un document vierge dans Word.</span><span class="sxs-lookup"><span data-stu-id="edbe9-120">Open Word and choose a blank document.</span></span> 
    
2. <span data-ttu-id="edbe9-121">Sous l’onglet **Insertion** , dans le groupe Compléments, cliquez sur **Store** et sélectionnez le complément **QR4Office**.</span><span class="sxs-lookup"><span data-stu-id="edbe9-121">On the Insert tab, in the Add-ins group, choose Store and select the QR4Office add-in. (You can load any add-in from the Store or your add-in catalog.)</span></span> <span data-ttu-id="edbe9-122">(Vous pouvez charger n’importe quel complément depuis le Store ou votre catalogue de compléments.)</span><span class="sxs-lookup"><span data-stu-id="edbe9-122">On the  Insert tab, in the Add-ins group, Store and select the QR4Office add-in. (You can load any add-in from the Store or your add-in catalog.)</span></span>
    
3. <span data-ttu-id="edbe9-123">Ouvrez les outils de développement F12 correspondant à votre version d’Office :</span><span class="sxs-lookup"><span data-stu-id="edbe9-123">Launch the F12 development tools that corresponds to your version of Office:</span></span>
    
   - <span data-ttu-id="edbe9-124">Pour la version 32 bits d'Office, utilisez C:\Windows\System32\F12\IEChooser.exe</span><span class="sxs-lookup"><span data-stu-id="edbe9-124">For the 32-bit version of Office, use C:\Windows\System32\F12\F12Chooser.exe</span></span>
    
   - <span data-ttu-id="edbe9-125">Pour la version 64 bits d'Office, utilisez C:\Windows\SysWOW64\F12\IEChooser.exe</span><span class="sxs-lookup"><span data-stu-id="edbe9-125">For the 64-bit version of Office, use C:\Windows\SysWOW64\F12\F12Chooser.exe</span></span>
    
   <span data-ttu-id="edbe9-126">Lorsque vouslancez F12Chooser, une autre fenêtre (intitulée « Choisir la cible à déboguer ») affiche les éventuelles applications à débogue.</span><span class="sxs-lookup"><span data-stu-id="edbe9-126">When you launch F12Chooser, a separate window named "Choose target to debug" displays the possible applications to debug.</span></span> <span data-ttu-id="edbe9-127">Sélectionnez l’application qui vous intéresse.</span><span class="sxs-lookup"><span data-stu-id="edbe9-127">Select the application that you are interested in.</span></span> <span data-ttu-id="edbe9-128">Si vous écrivez votre propre complément, sélectionnez le site web où le complément est déployé. Il peut s’agir d’une URL localhost.</span><span class="sxs-lookup"><span data-stu-id="edbe9-128">If you are writing your own add-in, select the website where you have the add-in deployed, which might be a localhost URL.</span></span> 
    
   <span data-ttu-id="edbe9-129">Par exemple, sélectionnez **home.html**.</span><span class="sxs-lookup"><span data-stu-id="edbe9-129">For example, select **home.html**.</span></span> 
    
   ![Écran IEChooser, pointant sur le complément bulles](../images/choose-target-to-debug.png)

4. <span data-ttu-id="edbe9-131">Dans la fenêtre F12, sélectionnez le fichier à déboguer.</span><span class="sxs-lookup"><span data-stu-id="edbe9-131">In the F12 window, select the file you want to debug.</span></span>
    
   <span data-ttu-id="edbe9-132">Pour sélectionner le fichier dans la fenêtre F12, cliquez sur l’icône de dossier située au-dessus du volet (gauche) du **script**.</span><span class="sxs-lookup"><span data-stu-id="edbe9-132">To select the file, choose the folder icon above the  **script** (left) pane.</span></span> <span data-ttu-id="edbe9-133">Dans la liste des fichiers disponibles affichés dans la liste déroulante, sélectionnez **Home.js**.</span><span class="sxs-lookup"><span data-stu-id="edbe9-133">From the list of available files shown in the dropdown list, select **Home.js**.</span></span>
    
5. <span data-ttu-id="edbe9-134">Définissez le point d’arrêt.</span><span class="sxs-lookup"><span data-stu-id="edbe9-134">Set the breakpoint.</span></span>
    
   <span data-ttu-id="edbe9-135">Pour définir le point d’arrêt dans **Home.js**, choisissez la ligne 144 située dans la fonction `textChanged`.</span><span class="sxs-lookup"><span data-stu-id="edbe9-135">To set the breakpoint in home.js, choose line 144, which is in the  textChanged function.</span></span> <span data-ttu-id="edbe9-136">Vous verrez un point rouge à gauche de la ligne et une ligne correspondante dans le volet **Pile d’appels et Points d’arrêt** (en bas à droite).</span><span class="sxs-lookup"><span data-stu-id="edbe9-136">You will see a red dot to the left of the line and a corresponding line in the **Callstack and Breakpoints** (bottom right) pane.</span></span> <span data-ttu-id="edbe9-137">Pour connaître d’autres façons de définir un point d’arrêt, consultez la rubrique [Inspecter le code JavaScript en cours d’exécution avec le débogueur](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span><span class="sxs-lookup"><span data-stu-id="edbe9-137">For other ways to set a breakpoint, see [Inspect running JavaScript with the Debugger](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span></span> 
    
   ![Débogueur avec le point d’arrêt dans le fichier home.js](../images/debugger-home-js-02.png)

6. <span data-ttu-id="edbe9-139">Exécutez votre complément pour déclencher le point d’arrêt.</span><span class="sxs-lookup"><span data-stu-id="edbe9-139">Run your add-in to trigger the breakpoint.</span></span>
    
   <span data-ttu-id="edbe9-140">Dans Word, cliquez sur la zone de texte URL dans la partie supérieure du volet **QR4Office** et essayez de saisir du texte.</span><span class="sxs-lookup"><span data-stu-id="edbe9-140">In Word, choose the URL textbox in the upper part of the **QR4Office** pane and attempt to enter some text.</span></span> <span data-ttu-id="edbe9-141">Dans le débogueur, dans le volet **Pile d’appels et Points d’arrêt**, vous verrez que le point d’arrêt s’est déclenché et affiche différentes informations.</span><span class="sxs-lookup"><span data-stu-id="edbe9-141">In the Debugger, in the  **Callstack and Breakpoints** pane, you'll see that the breakpoint has triggered and shows various information.</span></span> <span data-ttu-id="edbe9-142">Vous devrez peut-être actualiser le débogueur pour afficher les résultats.</span><span class="sxs-lookup"><span data-stu-id="edbe9-142">You might need to refresh the F12 tool to see the results.</span></span>
    
   ![Débogueur avec les résultats du point d’arrêt déclenché](../images/debugger-home-js-01.png)


## <a name="see-also"></a><span data-ttu-id="edbe9-144">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="edbe9-144">See also</span></span>

- <span data-ttu-id="edbe9-145">[Inspecter le code JavaScript en cours d’exécution avec le débogueur](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="edbe9-145">[Inspect running JavaScript with the Debugger](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span></span>
- <span data-ttu-id="edbe9-146">[Utilisation des outils de développement F12](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="edbe9-146">[Using the F12 developer tools](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span></span>
