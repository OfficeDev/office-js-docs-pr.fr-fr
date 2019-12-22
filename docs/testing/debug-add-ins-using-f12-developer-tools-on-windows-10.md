---
title: Débogage des compléments avec les outils de développement sur Windows 10
description: Débogage des compléments avec les outils de développement Microsoft Edge sur Windows 10
ms.date: 12/16/2019
localization_priority: Priority
ms.openlocfilehash: dd0f078b1cad0621dc99269b745a326adb69f7d6
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814471"
---
# <a name="debug-add-ins-using-developer-tools-on-windows-10"></a><span data-ttu-id="75b0a-103">Débogage des compléments avec les outils de développement sur Windows 10</span><span class="sxs-lookup"><span data-stu-id="75b0a-103">Debug add-ins using developer tools on Windows 10</span></span>

<span data-ttu-id="75b0a-104">Il existe des outils de développement en dehors des IDE pour vous aider à déboguer vos compléments sous Windows 10.</span><span class="sxs-lookup"><span data-stu-id="75b0a-104">There are developer tools outside of IDEs available to help you debug your add-ins on Windows 10.</span></span> <span data-ttu-id="75b0a-105">Ils sont utiles lorsque vous devez examiner un problème pendant l’exécution de votre complément hors de l’IDE.</span><span class="sxs-lookup"><span data-stu-id="75b0a-105">These are useful when you need to investigate a problem while running your add-in outside the IDE.</span></span>

<span data-ttu-id="75b0a-106">L’outil que vous utilisez dépend de l’exécution du complément dans Microsoft Edge ou Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="75b0a-106">The tool that you use depends on whether the add-in is running in Microsoft Edge or Internet Explorer.</span></span> <span data-ttu-id="75b0a-107">Cela est fonction de la version de Windows 10 et de la version d’Office qui sont installées sur l’ordinateur.</span><span class="sxs-lookup"><span data-stu-id="75b0a-107">This is determined by the version of Windows 10 and the version of Office that are installed on the computer.</span></span> <span data-ttu-id="75b0a-108">Pour déterminer quel navigateur est utilisé sur votre ordinateur de développement, consultez [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="75b0a-108">To determine which browser is being used on your development computer, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).</span></span>

> [!NOTE]
> <span data-ttu-id="75b0a-109">Les instructions décrites dans cet article ne peuvent pas être utilisées pour déboguer un complément Outlook qui utilise des fonctions Exécuter.</span><span class="sxs-lookup"><span data-stu-id="75b0a-109">The instructions in this article cannot be used to debug an Outlook add-in that uses Execute Functions.</span></span> <span data-ttu-id="75b0a-110">Pour déboguer un complément Outlook qui utilise des fonctions Exécuter, nous vous recommandons de l’attacher à Visual Studio en mode script ou à un autre débogueur de script.</span><span class="sxs-lookup"><span data-stu-id="75b0a-110">To debug an Outlook add-in that uses Execute Functions, we recommend that you attach to Visual Studio in script mode or to some other script debugger.</span></span>

## <a name="when-the-add-in-is-running-in-microsoft-edge"></a><span data-ttu-id="75b0a-111">Lorsque le complément s’exécute dans Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="75b0a-111">When the add-in is running in Microsoft Edge</span></span>

[!include[Enable debugging on Microsoft Edge DevTools](../includes/enable-debugging-on-edge-devtools.md)]

### <a name="debug-using-microsoft-edge-devtools"></a><span data-ttu-id="75b0a-112">Débogage avec Microsoft Edge DevTools</span><span class="sxs-lookup"><span data-stu-id="75b0a-112">Debug using Microsoft Edge DevTools</span></span>

<span data-ttu-id="75b0a-113">Lorsque le complément s’exécute dans Microsoft Edge, vous pouvez utiliser [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab).</span><span class="sxs-lookup"><span data-stu-id="75b0a-113">When the add-in is running in Microsoft Edge, you can use the [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab).</span></span>

1. <span data-ttu-id="75b0a-114">Exécutez le complément.</span><span class="sxs-lookup"><span data-stu-id="75b0a-114">Run the add-in.</span></span>

2. <span data-ttu-id="75b0a-115">Exécutez Microsoft Edge DevTools.</span><span class="sxs-lookup"><span data-stu-id="75b0a-115">Run the Microsoft Edge DevTools.</span></span>

3. <span data-ttu-id="75b0a-116">Dans les outils, ouvrez l’onglet **Local**. Votre complément est répertorié par son nom.</span><span class="sxs-lookup"><span data-stu-id="75b0a-116">In the tools, open the **Local** tab. Your add-in will be listed by its name.</span></span>

4. <span data-ttu-id="75b0a-117">Cliquez sur le nom du complément pour l’ouvrir dans les outils.</span><span class="sxs-lookup"><span data-stu-id="75b0a-117">Click the add-in name to open it in the tools.</span></span>

5. <span data-ttu-id="75b0a-118">Ouvrez l’onglet **Débogueur**.</span><span class="sxs-lookup"><span data-stu-id="75b0a-118">Open the **Debugger** tab.</span></span> 

6. <span data-ttu-id="75b0a-119">Cliquez sur l’icône de dossier située au-dessus du volet (gauche) du **script**.</span><span class="sxs-lookup"><span data-stu-id="75b0a-119">Choose the folder icon above the **script** (left) pane.</span></span> <span data-ttu-id="75b0a-120">Dans la liste des fichiers disponibles qui apparaît dans la liste déroulante, sélectionnez le fichier JavaScript que vous souhaitez déboguer.</span><span class="sxs-lookup"><span data-stu-id="75b0a-120">From the list of available files shown in the dropdown list, select the JavaScript file that you want to debug.</span></span>

7. <span data-ttu-id="75b0a-121">Pour définir un point d’arrêt, sélectionnez la ligne.</span><span class="sxs-lookup"><span data-stu-id="75b0a-121">To set a breakpoint, select the line.</span></span> <span data-ttu-id="75b0a-122">Vous verrez un point rouge à gauche de la ligne et une ligne correspondante dans le volet **Pile d’appels** (en bas à droite).</span><span class="sxs-lookup"><span data-stu-id="75b0a-122">You will see a red dot to the left of the line and a corresponding line in the **Call stack** (bottom right) pane.</span></span>

8. <span data-ttu-id="75b0a-123">Exécutez les fonctions dans le complément, si nécessaire, afin de déclencher le point d’arrêt.</span><span class="sxs-lookup"><span data-stu-id="75b0a-123">Execute functions in the add-in as needed to trigger the breakpoint.</span></span>

## <a name="when-the-add-in-is-running-in-internet-explorer"></a><span data-ttu-id="75b0a-124">Lorsque le complément s’exécute dans Internet Explorer</span><span class="sxs-lookup"><span data-stu-id="75b0a-124">When the add-in is running in Internet Explorer</span></span>

<span data-ttu-id="75b0a-125">Lorsque le complément s’exécute dans Internet Explorer, vous pouvez utiliser le débogueur des outils de développement F12 sous Windows 10 pour tester votre complément.</span><span class="sxs-lookup"><span data-stu-id="75b0a-125">When the add-in is running in Internet Explorer, you can use the debugger from the F12 developer tools in Windows 10 to test your add-in.</span></span> <span data-ttu-id="75b0a-126">Vous pouvez lancer les outils de développement F12 après l’exécution du complément.</span><span class="sxs-lookup"><span data-stu-id="75b0a-126">You can start the F12 developer tools after the add-in is running.</span></span> <span data-ttu-id="75b0a-127">Les outils F12 s’ouvrent dans une fenêtre distincte et n’utilisent pas Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="75b0a-127">The F12 tools are displayed in a separate window and do not use Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="75b0a-p107">Le débogueur fait partie des outils de développement F12 de Windows 10 et d’Internet Explorer. Il n’est pas inclus dans les versions antérieures de Windows.</span><span class="sxs-lookup"><span data-stu-id="75b0a-p107">The Debugger is part of the F12 developer tools in Windows 10 and Internet Explorer. Earlier versions of Windows do not include the Debugger.</span></span> 

<span data-ttu-id="75b0a-130">Cet exemple utilise Word et un complément gratuit d’AppSource.</span><span class="sxs-lookup"><span data-stu-id="75b0a-130">This example uses Word and a free add-in from AppSource.</span></span>

1. <span data-ttu-id="75b0a-131">Ouvrez un document vierge dans Word. </span><span class="sxs-lookup"><span data-stu-id="75b0a-131">Open Word and choose a blank document.</span></span> 
    
2. <span data-ttu-id="75b0a-132">Sous l’onglet **Insertion**, dans le groupe Compléments, cliquez sur **Store** et sélectionnez le complément **QR4Office**.</span><span class="sxs-lookup"><span data-stu-id="75b0a-132">On the **Insert** tab, in the Add-ins group, choose **Store** and select the **QR4Office** Add-in.</span></span> <span data-ttu-id="75b0a-133">(Vous pouvez charger n’importe quel complément depuis l’Office Store ou votre catalogue de compléments.)</span><span class="sxs-lookup"><span data-stu-id="75b0a-133">(You can load any add-in from the Store or your add-in catalog.)</span></span>
    
3. <span data-ttu-id="75b0a-134">Ouvrez les outils de développement F12 correspondant à votre version d’Office :</span><span class="sxs-lookup"><span data-stu-id="75b0a-134">Launch the F12 development tools that corresponds to your version of Office:</span></span>
    
   - <span data-ttu-id="75b0a-135">Pour la version 32 bits, utilisez C:\Windows\System32\F12\IEChooser.exe</span><span class="sxs-lookup"><span data-stu-id="75b0a-135">For the 32-bit version of Office, use C:\Windows\System32\F12\IEChooser.exe</span></span>
    
   - <span data-ttu-id="75b0a-136">Pour la version 64 bits, utilisez C:\Windows\SysWOW64\F12\IEChooser.exe</span><span class="sxs-lookup"><span data-stu-id="75b0a-136">For the 64-bit version of Office, use C:\Windows\SysWOW64\F12\IEChooser.exe</span></span>
    
   <span data-ttu-id="75b0a-137">Lorsque vous cliquez sur IEChooser, une autre fenêtre (intitulée « Choisir la cible à déboguer ») affiche les applications possibles pour effectuer le débogage.</span><span class="sxs-lookup"><span data-stu-id="75b0a-137">When you launch IEChooser, a separate window named "Choose target to debug" displays the possible applications to debug.</span></span> <span data-ttu-id="75b0a-138">Sélectionnez l’application de votre choix.</span><span class="sxs-lookup"><span data-stu-id="75b0a-138">Select the application that you are interested in.</span></span> <span data-ttu-id="75b0a-139">Si vous écrivez votre propre complément, sélectionnez le site web où le complément est déployé. Il peut s’agir d’une URL localhost.</span><span class="sxs-lookup"><span data-stu-id="75b0a-139">If you are writing your own add-in, select the website where you have the add-in deployed, which might be a localhost URL.</span></span> 
    
   <span data-ttu-id="75b0a-140">Par exemple, sélectionnez **home.html**.</span><span class="sxs-lookup"><span data-stu-id="75b0a-140">For example, select **home.html**.</span></span> 
    
   ![Écran IEChooser, pointant sur le complément bulles](../images/choose-target-to-debug.png)

4. <span data-ttu-id="75b0a-142">Dans la fenêtre F12, sélectionnez le fichier à déboguer.</span><span class="sxs-lookup"><span data-stu-id="75b0a-142">In the F12 window, select the file you want to debug.</span></span>
    
   <span data-ttu-id="75b0a-143">Pour sélectionner le fichier dans la fenêtre F12, cliquez sur l’icône de dossier située au-dessus du volet (gauche) du **script**.</span><span class="sxs-lookup"><span data-stu-id="75b0a-143">To select the file in the F12 window, choose the folder icon above the **script** (left) pane.</span></span> <span data-ttu-id="75b0a-144">Dans la liste des fichiers disponibles qui apparaît dans la liste déroulante, sélectionnez **Home.js**.</span><span class="sxs-lookup"><span data-stu-id="75b0a-144">From the list of available files shown in the dropdown list, select **Home.js**.</span></span>
    
5. <span data-ttu-id="75b0a-145">Définissez le point d’arrêt.</span><span class="sxs-lookup"><span data-stu-id="75b0a-145">Set the breakpoint.</span></span>
    
   <span data-ttu-id="75b0a-146">Pour définir le point d’arrêt dans **Home.js**, choisissez la ligne 144 située dans la fonction `textChanged`.</span><span class="sxs-lookup"><span data-stu-id="75b0a-146">To set the breakpoint in **Home.js**, choose line 144, which is in the  `textChanged` function.</span></span> <span data-ttu-id="75b0a-147">Vous verrez un point rouge à gauche de la ligne et une ligne correspondante dans le volet Pile d’appels et Points d’arrêt (en bas à droite).</span><span class="sxs-lookup"><span data-stu-id="75b0a-147">You will see a red dot to the left of the line and a corresponding line in the **Call stack and Breakpoints** (bottom right) pane.</span></span> <span data-ttu-id="75b0a-148">Pour connaître d’autres façons de définir un point d’arrêt, consultez la rubrique [Inspecter le code JavaScript en cours d’exécution avec le débogueur](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span><span class="sxs-lookup"><span data-stu-id="75b0a-148">For other ways to set a breakpoint, see [Inspect running JavaScript with the Debugger](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span></span> 
    
   ![Débogueur avec le point d’arrêt dans le fichier home.js](../images/debugger-home-js-02.png)

6. <span data-ttu-id="75b0a-150">Exécutez votre complément pour déclencher le point d’arrêt.</span><span class="sxs-lookup"><span data-stu-id="75b0a-150">Run your add-in to trigger the breakpoint.</span></span>
    
   <span data-ttu-id="75b0a-151">Dans Word, cliquez sur la zone de texte URL dans la partie supérieure du volet **QR4Office** et essayez de saisir du texte.</span><span class="sxs-lookup"><span data-stu-id="75b0a-151">In Word, choose the URL textbox in the upper part of the **QR4Office** pane and attempt to enter some text.</span></span> <span data-ttu-id="75b0a-152">Dans le débogueur, dans le volet **Pile d’appels et Points d’arrêt**, vous verrez que le point d’arrêt s’est déclenché et affiche différentes informations.</span><span class="sxs-lookup"><span data-stu-id="75b0a-152">In the Debugger, in the **Call stack and Breakpoints** pane, you'll see that the breakpoint has triggered and shows various information.</span></span> <span data-ttu-id="75b0a-153">Vous devrez peut-être actualiser le débogueur pour afficher les résultats.</span><span class="sxs-lookup"><span data-stu-id="75b0a-153">You might need to refresh the Debugger to see the results.</span></span>
    
   ![Débogueur avec les résultats du point d’arrêt déclenché](../images/debugger-home-js-01.png)


## <a name="see-also"></a><span data-ttu-id="75b0a-155">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="75b0a-155">See also</span></span>

- <span data-ttu-id="75b0a-156">[Inspecter le code JavaScript en cours d’exécution avec le débogueur](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="75b0a-156">[Inspect running JavaScript with the Debugger](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span></span>
- <span data-ttu-id="75b0a-157">[Utilisation des outils de développement F12](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="75b0a-157">[Using the F12 developer tools](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span></span>
