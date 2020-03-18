---
title: Déboguer des compléments Office sur un Mac
description: Découvrez comment utiliser un Mac pour déboguer des compléments Office
ms.date: 11/26/2019
localization_priority: Normal
ms.openlocfilehash: 0cd7edf8db40cbcb9057dc07e549e11e11b2c51c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719769"
---
# <a name="debug-office-add-ins-on-a-mac"></a><span data-ttu-id="64d44-103">Déboguer des compléments Office sur un Mac</span><span class="sxs-lookup"><span data-stu-id="64d44-103">Debug Office Add-ins on a Mac</span></span>

<span data-ttu-id="64d44-p101">Étant donné que les compléments sont développés avec du code HTML et JavaScript, ils sont conçus pour fonctionner sur toutes les plateformes, mais il peut y avoir de subtiles différences dans le rendu du code HTML par les différents navigateurs. Cet article décrit la procédure de débogage des compléments qui s’exécutent sur un Mac.</span><span class="sxs-lookup"><span data-stu-id="64d44-p101">Because add-ins are developed using HTML and JavaScript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on a Mac.</span></span>

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a><span data-ttu-id="64d44-106">Débogage avec l’inspecteur web Safari sur Mac</span><span class="sxs-lookup"><span data-stu-id="64d44-106">Debugging with Safari Web Inspector on a Mac</span></span>

<span data-ttu-id="64d44-107">Si votre complément affiche une interface utilisateur dans un volet des tâches ou dans un complément de contenu, vous pouvez déboguer un complément Office à l’aide de avec l’inspecteur web Safari.</span><span class="sxs-lookup"><span data-stu-id="64d44-107">If you have add-in that shows UI in a task pane or in a content add-in, you can debug an Office Add-in using Safari Web Inspector.</span></span>

<span data-ttu-id="64d44-108">Pour pouvoir déboguer des compléments Office sur Mac, vous devez disposer de Mac OS High Sierra ET de Mac Office version 16.9.1 (build 18012504) ou version ultérieure.</span><span class="sxs-lookup"><span data-stu-id="64d44-108">To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office Version: 16.9.1 (Build 18012504) or later.</span></span> <span data-ttu-id="64d44-109">Si vous n’avez pas de build Office Mac, vous pouvez en obtenir une en rejoignant le [Programme pour les développeurs Office 365](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="64d44-109">If you don't have an Office Mac build, you can get one by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).</span></span>

<span data-ttu-id="64d44-110">Pour commencer, ouvrez un terminal, puis définissez la propriété `OfficeWebAddinDeveloperExtras` pour l’application Office pertinente comme suit :</span><span class="sxs-lookup"><span data-stu-id="64d44-110">To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:</span></span>

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

<span data-ttu-id="64d44-111">Ensuite, ouvrez l’application Office et[insérez votre complément](sideload-an-office-add-in-on-ipad-and-mac.md).</span><span class="sxs-lookup"><span data-stu-id="64d44-111">Then, open the Office application and [sideload your add-in](sideload-an-office-add-in-on-ipad-and-mac.md).</span></span> <span data-ttu-id="64d44-112">Cliquez sur le complément. Vous devriez voir l’option **Inspecter l’élément** s’afficher dans le menu contextuel.</span><span class="sxs-lookup"><span data-stu-id="64d44-112">Right-click the add-in and you should see an **Inspect Element** option in the context menu.</span></span> <span data-ttu-id="64d44-113">Sélectionnez cette option pour afficher l’inspecteur dans lequel vous pouvez définir des points d’arrêt et déboguer votre complément.</span><span class="sxs-lookup"><span data-stu-id="64d44-113">Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="64d44-114">Si vous essayez d’utiliser l’inspecteur et si la boîte de dialogue scintille, mettez Office à jour vers la dernière version.</span><span class="sxs-lookup"><span data-stu-id="64d44-114">If you're trying to use the inspector and the dialog flickers, update Office to the latest version.</span></span> <span data-ttu-id="64d44-115">Si cela ne résout pas le problème de scintillement, essayez la solution de contournement suivante :</span><span class="sxs-lookup"><span data-stu-id="64d44-115">If that doesn't resolve the flickering, try the following workaround:</span></span>
> 1. <span data-ttu-id="64d44-116">Pour réduire la taille de la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="64d44-116">Reduce the size of the dialog.</span></span>
> 2. <span data-ttu-id="64d44-117">Sélectionnez l’option **Inspecter l’élément** qui ouvre une nouvelle fenêtre.</span><span class="sxs-lookup"><span data-stu-id="64d44-117">Choose **Inspect Element**, which opens in a new window.</span></span>
> 3. <span data-ttu-id="64d44-118">Redimensionner la boîte de dialogue à sa taille d’origine.</span><span class="sxs-lookup"><span data-stu-id="64d44-118">Resize the dialog to its original size.</span></span>
> 4. <span data-ttu-id="64d44-119">Utiliser l’inspecteur comme requis.</span><span class="sxs-lookup"><span data-stu-id="64d44-119">Use the inspector as required.</span></span>

## <a name="clearing-the-office-applications-cache-on-a-mac"></a><span data-ttu-id="64d44-120">Effacement du cache de l’application Office sur un ordinateur Mac</span><span class="sxs-lookup"><span data-stu-id="64d44-120">Clearing the Office application's cache on a Mac</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]
