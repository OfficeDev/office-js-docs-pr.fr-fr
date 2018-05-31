---
title: Débogage des compléments avec les outils de développement F12 sur Windows 10
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: e1e4cde4a1a0fe27058346b93e8aaa39dd75a4e3
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19438724"
---
# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a><span data-ttu-id="c41f6-102">Débogage des compléments avec les outils de développement F12 sur Windows 10</span><span class="sxs-lookup"><span data-stu-id="c41f6-102">Debug add-ins using F12 developer tools on Windows 10</span></span>

<span data-ttu-id="c41f6-p101">Les outils de développement F12 inclus dans Windows 10 vous aident à déboguer, tester et accélérer vos pages web. Ils vous aident également à développer et déboguer les compléments Office si vous n’utilisez pas un IDE comme Visual Studio ou si vous devez examiner un problème pendant l’exécution de votre complément hors de l’IDE. Vous pouvez lancer les outils de développement F12 après l’exécution de votre complément.</span><span class="sxs-lookup"><span data-stu-id="c41f6-p101">The F12 developer tools included in Windows 10 help you debug, test, and speed up your webpages. You can also use them to develop and debug Office Add-ins, if you are not using an IDE like Visual Studio, or if you need to investigate a problem while running your add-in outside the IDE. You can start the F12 developer tools after your add-in is running.</span></span>

<span data-ttu-id="c41f6-p102">Dans cet article, vous découvrirez comment utiliser le débogueur des outils de développement F12 de Windows 10 pour tester votre complément Office. Vous pouvez tester les compléments d’AppSource ou des compléments que vous avez ajoutés à partir d’autres emplacements. Les outils F12 s’ouvrent dans une fenêtre séparée et n’utilisent pas Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="c41f6-p102">This article shows how you how to use the Debugger tool from the F12 developer tools in Windows 10 to test your Office Add-in. You can test add-ins from AppSource or add-ins that you have added from other locations. The F12 tools display in a separate window and do not use Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="c41f6-p103">Le débogueur fait partie des outils de développement F12 de Windows 10 et d’Internet Explorer. Il n’est pas inclus dans les versions antérieures de Windows.</span><span class="sxs-lookup"><span data-stu-id="c41f6-p103">The Debugger is part of the F12 developer tools in Windows 10 and Internet Explorer. Earlier versions of Windows do not include the Debugger.</span></span> 

## <a name="prerequisites"></a><span data-ttu-id="c41f6-111">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="c41f6-111">Prerequisites</span></span>

<span data-ttu-id="c41f6-112">Les logiciels suivants doivent être installés :</span><span class="sxs-lookup"><span data-stu-id="c41f6-112">You need the following software:</span></span>

- <span data-ttu-id="c41f6-113">Les outils de développement F12, inclus dans Windows 10.</span><span class="sxs-lookup"><span data-stu-id="c41f6-113">The F12 developer tools, which are included in Windows 10.</span></span> 
    
- <span data-ttu-id="c41f6-114">L’application cliente Office qui héberge votre complément.</span><span class="sxs-lookup"><span data-stu-id="c41f6-114">The Office client application that hosts your add-in.</span></span> 
    
- <span data-ttu-id="c41f6-115">Votre complément.</span><span class="sxs-lookup"><span data-stu-id="c41f6-115">Your add-in.</span></span> 

## <a name="using-the-debugger"></a><span data-ttu-id="c41f6-116">Utilisation du débogueur</span><span class="sxs-lookup"><span data-stu-id="c41f6-116">Using the Debugger</span></span>

<span data-ttu-id="c41f6-117">Cet exemple utilise Word et un complément gratuit d’AppSource.</span><span class="sxs-lookup"><span data-stu-id="c41f6-117">This example uses Word and a free add-in from AppSource.</span></span>

1. <span data-ttu-id="c41f6-118">Ouvrez un document vierge dans Word.</span><span class="sxs-lookup"><span data-stu-id="c41f6-118">Open Word and choose a blank document.</span></span> 
    
2. <span data-ttu-id="c41f6-p104">Sous l’onglet **Insertion**, dans le groupe Compléments, cliquez sur **Store** et sélectionnez le complément QR4Office. (Vous pouvez charger n’importe quel complément depuis le Store ou votre catalogue de compléments.)</span><span class="sxs-lookup"><span data-stu-id="c41f6-p104">On the **Insert** tab, in the Add-ins group, choose **Store** and select the QR4Office add-in. (You can load any add-in from the Store or your add-in catalog.)</span></span>
    
3. <span data-ttu-id="c41f6-121">Ouvrez les outils de développement F12 correspondant à votre version d’Office :</span><span class="sxs-lookup"><span data-stu-id="c41f6-121">Launch the F12 development tools that corresponds to your version of Office:</span></span>
    
   - <span data-ttu-id="c41f6-122">Pour la version 32 bits, utilisez C:\Windows\System32\F12\F12Chooser.exe</span><span class="sxs-lookup"><span data-stu-id="c41f6-122">For the 32-bit version of Office, use C:\Windows\System32\F12\F12Chooser.exe</span></span>
    
   - <span data-ttu-id="c41f6-123">Pour la version 64 bits, utilisez C:\Windows\SysWOW64\F12\F12Chooser.exe</span><span class="sxs-lookup"><span data-stu-id="c41f6-123">For the 64-bit version of Office, use C:\Windows\SysWOW64\F12\F12Chooser.exe</span></span>
    
   <span data-ttu-id="c41f6-p105">Lorsque vous cliquez sur F12Chooser, une autre fenêtre (intitulée « Choisir la cible à déboguer ») affiche les applications possibles pour effectuer le débogage. Sélectionnez l’application de votre choix. Si vous écrivez votre propre complément, sélectionnez le site web où le complément est déployé. Il peut s’agir d’une URL localhost.</span><span class="sxs-lookup"><span data-stu-id="c41f6-p105">When you launch F12Chooser, a separate window named "Choose target to debug" displays the possible applications to debug. Select the application that you are interested in. If you are writing your own add-in, select the website where you have the add-in deployed, which might be a localhost URL.</span></span> 
    
   <span data-ttu-id="c41f6-127">Par exemple, sélectionnez **home.html**.</span><span class="sxs-lookup"><span data-stu-id="c41f6-127">For example, select **home.html**.</span></span> 
    
   ![Écran du sélecteur F12, pointe vers un complément de type « bulles »](../images/choose-target-to-debug.png)

4. <span data-ttu-id="c41f6-129">Dans la fenêtre F12, sélectionnez le fichier à déboguer.</span><span class="sxs-lookup"><span data-stu-id="c41f6-129">In the F12 window, select the file you want to debug.</span></span>
    
   <span data-ttu-id="c41f6-p106">Pour sélectionner le fichier, cliquez sur l’icône de dossier située au-dessus du volet (gauche) du **script**. La liste déroulante affiche les fichiers disponibles. Sélectionnez home.js.</span><span class="sxs-lookup"><span data-stu-id="c41f6-p106">To select the file, choose the folder icon above the  **script** (left) pane. The dropdown list shows the available files. Select home.js.</span></span>
    
5. <span data-ttu-id="c41f6-133">Définissez le point d’arrêt.</span><span class="sxs-lookup"><span data-stu-id="c41f6-133">Set the breakpoint.</span></span>
    
   <span data-ttu-id="c41f6-p107">Pour définir un point d'arrêt dans home.js, choisissez la ligne 144 qui se trouve dans la fonction _textChanged_. Vous verrez un point rouge à gauche de la ligne et une ligne correspondante dans le volet **Callstack and Breakpoints** (en bas à droite). Pour connaître d'autres manières de définir un point d'arrêt, référez-vous à [Consulter JavaScript en fonctionnement avec le débogueur](https://msdn.microsoft.com/library/dn255007%28v=vs.85%29.aspx).</span><span class="sxs-lookup"><span data-stu-id="c41f6-p107">To set the breakpoint in home.js, choose line 144, which is in the  _textChanged_ function. You will see a red dot to the left of the line and a corresponding line in the **Callstack and Breakpoints** (bottom right) pane. For other ways to set a breakpoint, see [Inspect running JavaScript with the Debugger](https://msdn.microsoft.com/library/dn255007%28v=vs.85%29.aspx).</span></span> 
    
   ![Débogueur avec le point d’arrêt dans le fichier home.js](../images/debugger-home-js-02.png)

6. <span data-ttu-id="c41f6-138">Exécutez votre complément pour déclencher le point d’arrêt.</span><span class="sxs-lookup"><span data-stu-id="c41f6-138">Run your add-in to trigger the breakpoint.</span></span>
    
   <span data-ttu-id="c41f6-p108">Cliquez sur la zone de texte URL dans la partie supérieure du volet QR4Office pour modifier le texte. Dans le débogueur, dans le volet **Pile d’appels et Points d’arrêt**, vous verrez que le point d’arrêt s’est déclenché et affiche différentes informations. Vous devrez peut-être actualiser l’outil F12 pour afficher les résultats.</span><span class="sxs-lookup"><span data-stu-id="c41f6-p108">Choose the URL textbox in the upper part of the QR4Office pane to change the text. In the Debugger, in the **Callstack and Breakpoints** pane, you'll see that the breakpoint has triggered and shows various information. You might need to refresh the F12 tool to see the results.</span></span>
    
   ![Débogueur avec les résultats du point d’arrêt déclenché](../images/debugger-home-js-01.png)


## <a name="see-also"></a><span data-ttu-id="c41f6-143">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c41f6-143">See also</span></span>

- [<span data-ttu-id="c41f6-144">Inspecter le code JavaScript en cours d’exécution avec le débogueur</span><span class="sxs-lookup"><span data-stu-id="c41f6-144">Inspect running JavaScript with the Debugger</span></span>](https://msdn.microsoft.com/library/dn255007%28v=vs.85%29.aspx)
- [<span data-ttu-id="c41f6-145">Utilisation des outils de développement F12</span><span class="sxs-lookup"><span data-stu-id="c41f6-145">Using the F12 developer tools</span></span>](https://msdn.microsoft.com/en-us/library/bg182326%28v=vs.85%29.aspx)
    
