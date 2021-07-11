---
title: Déboguer des compléments Office sur un Mac
description: Découvrez comment utiliser un Mac pour déboguer des Office des macros.
ms.date: 10/16/2020
localization_priority: Normal
ms.openlocfilehash: 98473e7c37b9ef5ee34d35f91688ccef65ac7d78
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350133"
---
# <a name="debug-office-add-ins-on-a-mac"></a><span data-ttu-id="c62c7-103">Déboguer des compléments Office sur un Mac</span><span class="sxs-lookup"><span data-stu-id="c62c7-103">Debug Office Add-ins on a Mac</span></span>

<span data-ttu-id="c62c7-p101">Étant donné que les compléments sont développés avec du code HTML et JavaScript, ils sont conçus pour fonctionner sur toutes les plateformes, mais il peut y avoir de subtiles différences dans le rendu du code HTML par les différents navigateurs. Cet article décrit la procédure de débogage des compléments qui s’exécutent sur un Mac.</span><span class="sxs-lookup"><span data-stu-id="c62c7-p101">Because add-ins are developed using HTML and JavaScript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on a Mac.</span></span>

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a><span data-ttu-id="c62c7-106">Débogage avec l’inspecteur web Safari sur Mac</span><span class="sxs-lookup"><span data-stu-id="c62c7-106">Debugging with Safari Web Inspector on a Mac</span></span>

<span data-ttu-id="c62c7-107">Si votre complément affiche une interface utilisateur dans un volet des tâches ou dans un complément de contenu, vous pouvez déboguer un complément Office à l’aide de avec l’inspecteur web Safari.</span><span class="sxs-lookup"><span data-stu-id="c62c7-107">If you have add-in that shows UI in a task pane or in a content add-in, you can debug an Office Add-in using Safari Web Inspector.</span></span>

<span data-ttu-id="c62c7-108">Pour pouvoir déboguer des Office sur Mac, vous devez avoir Mac OS High Sierra et Mac Office version 16.9.1 (build 18012504) ou version ultérieure.</span><span class="sxs-lookup"><span data-stu-id="c62c7-108">To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office version 16.9.1 (build 18012504) or later.</span></span> <span data-ttu-id="c62c7-109">Si vous n’avez pas de build Office Mac, vous pouvez en obtenir une en rejoignant le programme [Microsoft 365 développeur.](https://developer.microsoft.com/office/dev-program)</span><span class="sxs-lookup"><span data-stu-id="c62c7-109">If you don't have an Office Mac build, you can get one by joining the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

<span data-ttu-id="c62c7-110">Pour commencer, ouvrez un terminal, puis définissez la propriété `OfficeWebAddinDeveloperExtras` pour l’application Office pertinente comme suit :</span><span class="sxs-lookup"><span data-stu-id="c62c7-110">To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:</span></span>

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

    > [!IMPORTANT]
    > <span data-ttu-id="c62c7-111">Les builds d’applications du Mac App Store Office ne pas la prise en charge de `OfficeWebAddinDeveloperExtras` l’indicateur.</span><span class="sxs-lookup"><span data-stu-id="c62c7-111">Mac App Store builds of Office do not support the `OfficeWebAddinDeveloperExtras` flag.</span></span>

<span data-ttu-id="c62c7-112">Ensuite, ouvrez l’application Office et[insérez votre complément](sideload-an-office-add-in-on-ipad-and-mac.md).</span><span class="sxs-lookup"><span data-stu-id="c62c7-112">Then, open the Office application and [sideload your add-in](sideload-an-office-add-in-on-ipad-and-mac.md).</span></span> <span data-ttu-id="c62c7-113">Cliquez sur le complément. Vous devriez voir l’option **Inspecter l’élément** s’afficher dans le menu contextuel.</span><span class="sxs-lookup"><span data-stu-id="c62c7-113">Right-click the add-in and you should see an **Inspect Element** option in the context menu.</span></span> <span data-ttu-id="c62c7-114">Sélectionnez cette option pour afficher l’inspecteur dans lequel vous pouvez définir des points d’arrêt et déboguer votre complément.</span><span class="sxs-lookup"><span data-stu-id="c62c7-114">Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="c62c7-115">Si vous essayez d’utiliser l’inspecteur et si la boîte de dialogue scintille, mettez Office à jour vers la dernière version.</span><span class="sxs-lookup"><span data-stu-id="c62c7-115">If you're trying to use the inspector and the dialog flickers, update Office to the latest version.</span></span> <span data-ttu-id="c62c7-116">Si cela ne résout pas le clignotement, essayez la solution de contournement suivante.</span><span class="sxs-lookup"><span data-stu-id="c62c7-116">If that doesn't resolve the flickering, try the following workaround.</span></span>
>
> 1. <span data-ttu-id="c62c7-117">Pour réduire la taille de la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="c62c7-117">Reduce the size of the dialog.</span></span>
> 1. <span data-ttu-id="c62c7-118">Sélectionnez l’option **Inspecter l’élément** qui ouvre une nouvelle fenêtre.</span><span class="sxs-lookup"><span data-stu-id="c62c7-118">Choose **Inspect Element**, which opens in a new window.</span></span>
> 1. <span data-ttu-id="c62c7-119">Redimensionner la boîte de dialogue à sa taille d’origine.</span><span class="sxs-lookup"><span data-stu-id="c62c7-119">Resize the dialog to its original size.</span></span>
> 1. <span data-ttu-id="c62c7-120">Utiliser l’inspecteur comme requis.</span><span class="sxs-lookup"><span data-stu-id="c62c7-120">Use the inspector as required.</span></span>

## <a name="clearing-the-office-applications-cache-on-a-mac"></a><span data-ttu-id="c62c7-121">Effacement du cache de l’application Office sur un ordinateur Mac</span><span class="sxs-lookup"><span data-stu-id="c62c7-121">Clearing the Office application's cache on a Mac</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]
