---
title: Chargement de version test des compl?ments Office dans Office Online
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 823821f990674a2d822508a860a7e5d6424e0245
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="sideload-office-add-ins-in-office-online-for-testing"></a><span data-ttu-id="52169-102">Chargement de version test des compl?ments Office dans Office Online</span><span class="sxs-lookup"><span data-stu-id="52169-102">Sideload Office Add-ins in Office Online for testing</span></span>

<span data-ttu-id="52169-p101">Vous pouvez installer un compl?ment Office test sans avoir ? le placer au pr?alable dans un catalogue de compl?ments en utilisant le chargement de version test. Le chargement de version test peut ?tre effectu? sur Office 365 ou Office Online. La proc?dure pr?sente de l?g?res diff?rences d?une plateforme ? l?autre.</span><span class="sxs-lookup"><span data-stu-id="52169-p101">You can install an Office Add-in for testing without having to first put it in an add-in catalog by using sideloading. Sideloading can be done on either Office 365 or Office Online. The procedure is slightly different for the two platforms.</span></span> 

<span data-ttu-id="52169-106">Lorsque vous chargez une version test d?un compl?ment, le manifeste du compl?ment est stock? dans le stockage local du navigateur. Ainsi, si vous videz le cache du navigateur ou si vous basculez vers un autre navigateur, vous devez ? nouveau charger une version test de compl?ment.</span><span class="sxs-lookup"><span data-stu-id="52169-106">When you sideload an add-in, the add-in manifest is stored in the browser's local storage, so if you clear the browser's cache, or switch to a different browser, you have to sideload the add-in again.</span></span>


> [!NOTE]
> <span data-ttu-id="52169-p102">Tel que d?crit dans cet article, le chargement de version test est pris en charge dans Word, Excel et PowerPoint. Pour charger une version test de compl?ment Outlook, voir la rubrique relative au [chargement de version test des compl?ments Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span><span class="sxs-lookup"><span data-stu-id="52169-p102">Sideloading as described in this article is supported on Word, Excel, and PowerPoint. To sideload an Outlook add-in, see [Sideload Outlook add-ins for testing](https://docs.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span></span>

<span data-ttu-id="52169-109">La vid?o suivante pr?sente la proc?dure de chargement de version test de votre compl?ment dans la version de bureau Office ou Office Online.</span><span class="sxs-lookup"><span data-stu-id="52169-109">The following video walks you through the process of sideloading your add-in on Office desktop or Office Online.</span></span>  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="sideload-an-office-add-in-on-office-365"></a><span data-ttu-id="52169-110">Chargement de version test d?un compl?ment Office dans Office 365</span><span class="sxs-lookup"><span data-stu-id="52169-110">Sideload an Office Add-in on Office 365</span></span>


1. <span data-ttu-id="52169-111">Connectez-vous ? votre compte Office 365.</span><span class="sxs-lookup"><span data-stu-id="52169-111">Sign in to your Office 365 account.</span></span>
    
2. <span data-ttu-id="52169-112">Ouvrez le lanceur d?applications ? l?extr?mit? gauche de la barre d?outils et s?lectionnez **Excel**,  **Word** ou **PowerPoint**, puis cr?ez un document.</span><span class="sxs-lookup"><span data-stu-id="52169-112">Open the App Launcher on the left end of the toolbar and select  **Excel**,  **Word**, or  **PowerPoint**, and then create a new document.</span></span>
    
3. <span data-ttu-id="52169-113">Ouvrez l?onglet **Ins?rer** dans le ruban, puis dans la section **Compl?ments**, choisissez **Compl?ments Office**.</span><span class="sxs-lookup"><span data-stu-id="52169-113">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
    
4. <span data-ttu-id="52169-114">Dans la bo?te de dialogue **Compl?ments Office**, s?lectionnez l?onglet **MON ORGANISATION**, puis **T?l?charger mon compl?ment**.</span><span class="sxs-lookup"><span data-stu-id="52169-114">On the  **Office Add-ins** dialog, select the **MY ORGANIZATION** tab, and then **Upload My Add-in**.</span></span>
    
    ![Bo?te de dialogue intitul?e Compl?ment Office avec un lien dans le coin sup?rieur gauche indiquant ? Charger mon compl?ment ?.](../images/office-add-ins.png)

5.  <span data-ttu-id="52169-116">**Acc?dez** au fichier manifeste du compl?ment, puis s?lectionnez **T?l?charger**.</span><span class="sxs-lookup"><span data-stu-id="52169-116">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![Bo?te de dialogue de chargement de compl?ment avec des boutons pour parcourir, t?l?charger et annuler.](../images/upload-add-in.png)

6. <span data-ttu-id="52169-p103">Verify that your compl?ment is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in the pane should appear.</span><span class="sxs-lookup"><span data-stu-id="52169-p103">Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in the pane should appear.</span></span>
    

## <a name="sideload-an-office-add-in-on-office-online"></a><span data-ttu-id="52169-121">Charger une version test d?un compl?ment Office sur Office Online</span><span class="sxs-lookup"><span data-stu-id="52169-121">Sideload an Office Add-in on Office Online</span></span>


1. <span data-ttu-id="52169-122">Ouvrez [Microsoft Office Online](https://office.live.com/).</span><span class="sxs-lookup"><span data-stu-id="52169-122">Open [Microsoft Office Online](https://office.live.com/).</span></span>
    
2. <span data-ttu-id="52169-123">Dans **Commencer ? utiliser les applications en ligne maintenant**, choisissez **Excel**, **Word** ou **PowerPoint**, puis ouvrez un document.</span><span class="sxs-lookup"><span data-stu-id="52169-123">In  **Get started with the online apps now**, choose  **Excel**,  **Word**, or  **PowerPoint**; and then open a new document.</span></span>
    
3. <span data-ttu-id="52169-124">Ouvrez l?onglet **Ins?rer** dans le ruban, puis dans la section **Compl?ments**, choisissez **Compl?ments Office**.</span><span class="sxs-lookup"><span data-stu-id="52169-124">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
    
4. <span data-ttu-id="52169-125">Dans la bo?te de dialogue **Compl?ments Office**, s?lectionnez l?onglet **MES COMPL?MENTS**, choisissez **G?rer mes compl?ments**, puis **T?l?charger mon compl?ment**.</span><span class="sxs-lookup"><span data-stu-id="52169-125">On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.</span></span>
    
    ![Bo?te de dialogue Compl?ments Office avec une liste d?roulante dans le coin sup?rieur droit indiquant ? G?rer mes compl?ments ? et une autre liste d?roulante sous cette derni?re avec l?option ? Charger mon compl?ment ?](../images/office-add-ins-my-account.png)

5.  <span data-ttu-id="52169-127">**Acc?dez** au fichier manifeste du compl?ment, puis s?lectionnez **T?l?charger**.</span><span class="sxs-lookup"><span data-stu-id="52169-127">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![Bo?te de dialogue de t?l?chargement de compl?ment avec des boutons pour parcourir, t?l?charger et annuler.](../images/upload-add-in.png)

6. <span data-ttu-id="52169-p104">V?rifiez que votre compl?ment est install?. S?il s?agit d?une commande de compl?ment, elle doit appara?tre dans le ruban ou dans le menu contextuel. S?il s?agit d?un compl?ment du volet Office, le volet doit appara?tre.</span><span class="sxs-lookup"><span data-stu-id="52169-p104">Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in, the pane should appear.</span></span>

## <a name="sideload-an-add-in-when-using-visual-studio"></a><span data-ttu-id="52169-132">Chargement d?une version test d?un compl?ment lors de l?utilisation de Visual Studio</span><span class="sxs-lookup"><span data-stu-id="52169-132">Sideload an add-in when using Visual Studio</span></span>

<span data-ttu-id="52169-p105">Si vous d?veloppez votre compl?ment ? l?aide de Visual Studio, le processus de chargement d?une version de teste est similaire. La seule diff?rence est que vous devez mettre ? jour la valeur de l??l?ment **SourceURL** dans votre manifeste, de sorte ? inclure l?URL enti?re de l?emplacement de d?ploiement du compl?ment.</span><span class="sxs-lookup"><span data-stu-id="52169-p105">If you're using Visual Studio to develop your add-in, the process to sideload is similar. The only difference is that you will have to update the value of the **SourceURL** element in your manifest to include the full URL where the add-in is deployed.</span></span> 

<span data-ttu-id="52169-p106">Si vous ?tes en train de d?velopper votre compl?ment, recherchez-le dans le fichier manifest.xml et mettez ? jour la valeur de l??l?ment **SourceLocation** de fa?on ? inclure un URI absolu. Visual Studio met en place un jeton pour votre d?ploiement localhost.</span><span class="sxs-lookup"><span data-stu-id="52169-p106">If you're currently developing your add-in, locate your add-in manifest.xml file, and update the **SourceLocation** element value to include an absolute URI. Visual Studio will put in a token for your localhost deployment.</span></span>

<span data-ttu-id="52169-137">Par exemple :</span><span class="sxs-lookup"><span data-stu-id="52169-137">For example:</span></span> 

```xml
<SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
```
