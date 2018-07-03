---
title: Chargement de version test des compléments Office dans Office Online
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 69b255545525ff667618c9f8bd1e1b7953592967
ms.sourcegitcommit: 58af795c3d0393a4b1f6425fa1cbdca1e48fb473
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/29/2018
ms.locfileid: "20138848"
---
# <a name="sideload-office-add-ins-in-office-online-for-testing"></a><span data-ttu-id="05b2e-102">Chargement de version test des compléments Office dans Office Online</span><span class="sxs-lookup"><span data-stu-id="05b2e-102">Sideload Office Add-ins in Office Online for testing</span></span>

<span data-ttu-id="05b2e-p101">Vous pouvez installer un complément Office test sans avoir à le placer au préalable dans un catalogue de compléments en utilisant le chargement de version test. Le chargement de version test peut être effectué sur Office 365 ou Office Online. La procédure présente de légères différences d’une plateforme à l’autre.</span><span class="sxs-lookup"><span data-stu-id="05b2e-p101">You can install an Office Add-in for testing without having to first put it in an add-in catalog by using sideloading. Sideloading can be done on either Office 365 or Office Online. The procedure is slightly different for the two platforms.</span></span> 

<span data-ttu-id="05b2e-106">Lorsque vous chargez une version test d’un complément, le manifeste du complément est stocké dans le stockage local du navigateur. Ainsi, si vous videz le cache du navigateur ou si vous basculez vers un autre navigateur, vous devez à nouveau charger une version test de complément.</span><span class="sxs-lookup"><span data-stu-id="05b2e-106">When you sideload an add-in, the add-in manifest is stored in the browser's local storage, so if you clear the browser's cache, or switch to a different browser, you have to sideload the add-in again.</span></span>


> [!NOTE]
> <span data-ttu-id="05b2e-p102">Tel que décrit dans cet article, le chargement de version test est pris en charge dans Word, Excel et PowerPoint. Pour charger une version test de complément Outlook, voir la rubrique relative au [chargement de version test des compléments Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span><span class="sxs-lookup"><span data-stu-id="05b2e-p102">Sideloading as described in this article is supported on Word, Excel, and PowerPoint. To sideload an Outlook add-in, see [Sideload Outlook add-ins for testing](https://docs.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span></span>

<span data-ttu-id="05b2e-109">La vidéo suivante présente la procédure de chargement de version test de votre complément dans la version de bureau Office ou Office Online.</span><span class="sxs-lookup"><span data-stu-id="05b2e-109">The following video walks you through the process of sideloading your add-in on Office desktop or Office Online.</span></span>  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="sideload-an-office-add-in-on-office-365"></a><span data-ttu-id="05b2e-110">Chargement de version test d’un complément Office dans Office 365</span><span class="sxs-lookup"><span data-stu-id="05b2e-110">Sideload an Office Add-in on Office 365</span></span>


1. <span data-ttu-id="05b2e-111">Connectez-vous à votre compte Office 365.</span><span class="sxs-lookup"><span data-stu-id="05b2e-111">Sign in to your Office 365 account.</span></span>
    
2. <span data-ttu-id="05b2e-112">Ouvrez le lanceur d’applications à l’extrémité gauche de la barre d’outils et sélectionnez **Excel**,  **Word** ou **PowerPoint**, puis créez un document.</span><span class="sxs-lookup"><span data-stu-id="05b2e-112">Open the App Launcher on the left end of the toolbar and select  **Excel**,  **Word**, or  **PowerPoint**, and then create a new document.</span></span>
    
3. <span data-ttu-id="05b2e-113">Ouvrez l’onglet **Insérer** dans le ruban, puis dans la section **Compléments**, choisissez **Compléments Office**.</span><span class="sxs-lookup"><span data-stu-id="05b2e-113">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
    
4. <span data-ttu-id="05b2e-114">Dans la boîte de dialogue **Compléments Office**, sélectionnez l’onglet **MON ORGANISATION**, puis **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="05b2e-114">On the  **Office Add-ins** dialog, select the **MY ORGANIZATION** tab, and then **Upload My Add-in**.</span></span>
    
    ![Boîte de dialogue intitulée Complément Office avec un lien dans le coin supérieur gauche indiquant « Charger mon complément ».](../images/office-add-ins.png)

5.  <span data-ttu-id="05b2e-116">**Accédez** au fichier manifeste du complément, puis sélectionnez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="05b2e-116">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![Boîte de dialogue de chargement de complément avec des boutons pour parcourir, télécharger et annuler.](../images/upload-add-in.png)

6. <span data-ttu-id="05b2e-p103">Verify that your complément is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in the pane should appear.</span><span class="sxs-lookup"><span data-stu-id="05b2e-p103">Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in the pane should appear.</span></span>
    

## <a name="sideload-an-office-add-in-on-office-online"></a><span data-ttu-id="05b2e-121">Charger une version test d’un complément Office sur Office Online</span><span class="sxs-lookup"><span data-stu-id="05b2e-121">Sideload an Office Add-in on Office Online</span></span>


1. <span data-ttu-id="05b2e-122">Ouvrez [Microsoft Office Online](https://office.live.com/).</span><span class="sxs-lookup"><span data-stu-id="05b2e-122">Open [Microsoft Office Online](https://office.live.com/).</span></span>
    
2. <span data-ttu-id="05b2e-123">Dans **Commencer à utiliser les applications en ligne maintenant**, choisissez **Excel**, **Word** ou **PowerPoint**, puis ouvrez un document.</span><span class="sxs-lookup"><span data-stu-id="05b2e-123">In  **Get started with the online apps now**, choose  **Excel**,  **Word**, or  **PowerPoint**; and then open a new document.</span></span>
    
3. <span data-ttu-id="05b2e-124">Ouvrez l’onglet **Insérer** dans le ruban, puis dans la section **Compléments**, choisissez **Compléments Office**.</span><span class="sxs-lookup"><span data-stu-id="05b2e-124">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
    
4. <span data-ttu-id="05b2e-125">Dans la boîte de dialogue **Compléments Office**, sélectionnez l’onglet **MES COMPLÉMENTS**, choisissez **Gérer mes compléments**, puis **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="05b2e-125">On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.</span></span>
    
    ![Boîte de dialogue Compléments Office avec une liste déroulante dans le coin supérieur droit indiquant « Gérer mes compléments » et une autre liste déroulante sous cette dernière avec l’option « Charger mon complément »](../images/office-add-ins-my-account.png)

5.  <span data-ttu-id="05b2e-127">**Accédez** au fichier manifeste du complément, puis sélectionnez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="05b2e-127">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![Boîte de dialogue de téléchargement de complément avec des boutons pour parcourir, télécharger et annuler.](../images/upload-add-in.png)

6. <span data-ttu-id="05b2e-p104">Vérifiez que votre complément est installé. S’il s’agit d’une commande de complément, elle doit apparaître dans le ruban ou dans le menu contextuel. S’il s’agit d’un complément du volet Office, le volet doit apparaître.</span><span class="sxs-lookup"><span data-stu-id="05b2e-p104">Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in, the pane should appear.</span></span>

> [!NOTE]
><span data-ttu-id="05b2e-132">Pour tester votre complément Office avec Edge, entrez « **about:flags**   » dans la barre de recherche Edge pour afficher les options des paramètres de développement.</span><span class="sxs-lookup"><span data-stu-id="05b2e-132">To test your Office Add-in with Edge, enter “**about:flags**” in the Edge search bar to bring up the Developer Settings options.</span></span>  <span data-ttu-id="05b2e-133">Vérifiez l'option « **Autoriser le bouclage localhost** » et redémarrez Edge.</span><span class="sxs-lookup"><span data-stu-id="05b2e-133">Check the “**Allow localhost loopback**” option and restart Edge.</span></span>

>    ![L'option d'Edge « Autoriser le bouclage localhost » avec la case cochée.](../images/allow-localhost-loopback.png)

## <a name="sideload-an-add-in-when-using-visual-studio"></a><span data-ttu-id="05b2e-135">Chargement d’une version test d’un complément lors de l’utilisation de Visual Studio</span><span class="sxs-lookup"><span data-stu-id="05b2e-135">Sideload an add-in when using Visual Studio</span></span>

<span data-ttu-id="05b2e-p106">Si vous développez votre complément à l’aide de Visual Studio, le processus de chargement d’une version de teste est similaire. La seule différence est que vous devez mettre à jour la valeur de l’élément **SourceURL** dans votre manifeste, de sorte à inclure l’URL entière de l’emplacement de déploiement du complément.</span><span class="sxs-lookup"><span data-stu-id="05b2e-p106">If you're using Visual Studio to develop your add-in, the process to sideload is similar. The only difference is that you will have to update the value of the **SourceURL** element in your manifest to include the full URL where the add-in is deployed.</span></span> 

<span data-ttu-id="05b2e-p107">Si vous êtes en train de développer votre complément, recherchez-le dans le fichier manifest.xml et mettez à jour la valeur de l’élément **SourceLocation** de façon à inclure un URI absolu. Visual Studio met en place un jeton pour votre déploiement localhost.</span><span class="sxs-lookup"><span data-stu-id="05b2e-p107">If you're currently developing your add-in, locate your add-in manifest.xml file, and update the **SourceLocation** element value to include an absolute URI. Visual Studio will put in a token for your localhost deployment.</span></span>

<span data-ttu-id="05b2e-140">Par exemple :</span><span class="sxs-lookup"><span data-stu-id="05b2e-140">For example:</span></span> 

```xml
<SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
```
