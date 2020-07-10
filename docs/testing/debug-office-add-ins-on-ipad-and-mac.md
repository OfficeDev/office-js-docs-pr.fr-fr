---
title: Déboguer des compléments Office sur un Mac
description: Découvrez comment utiliser un Mac pour déboguer des compléments Office
ms.date: 11/26/2019
localization_priority: Normal
ms.openlocfilehash: 12785a195c336e0de8c619379a3839bd15079b2c
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094126"
---
# <a name="debug-office-add-ins-on-a-mac"></a><span data-ttu-id="1cf74-103">Déboguer des compléments Office sur un Mac</span><span class="sxs-lookup"><span data-stu-id="1cf74-103">Debug Office Add-ins on a Mac</span></span>

<span data-ttu-id="1cf74-104">Because add-ins are developed using HTML and JavaScript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML.</span><span class="sxs-lookup"><span data-stu-id="1cf74-104">Because add-ins are developed using HTML and JavaScript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML.</span></span> <span data-ttu-id="1cf74-105">This article describes how to debug add-ins running on a Mac.</span><span class="sxs-lookup"><span data-stu-id="1cf74-105">This article describes how to debug add-ins running on a Mac.</span></span>

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a><span data-ttu-id="1cf74-106">Débogage avec l’inspecteur web Safari sur Mac</span><span class="sxs-lookup"><span data-stu-id="1cf74-106">Debugging with Safari Web Inspector on a Mac</span></span>

<span data-ttu-id="1cf74-107">Si votre complément affiche une interface utilisateur dans un volet des tâches ou dans un complément de contenu, vous pouvez déboguer un complément Office à l’aide de avec l’inspecteur web Safari.</span><span class="sxs-lookup"><span data-stu-id="1cf74-107">If you have add-in that shows UI in a task pane or in a content add-in, you can debug an Office Add-in using Safari Web Inspector.</span></span>

<span data-ttu-id="1cf74-108">Pour pouvoir déboguer des compléments Office sur Mac, vous devez disposer de Mac OS High Sierra ET de Mac Office version 16.9.1 (build 18012504) ou version ultérieure.</span><span class="sxs-lookup"><span data-stu-id="1cf74-108">To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office Version: 16.9.1 (Build 18012504) or later.</span></span> <span data-ttu-id="1cf74-109">Si vous n’avez pas de version Office Mac, vous pouvez en obtenir une en rejoignant le programme pour les [développeurs Microsoft 365](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="1cf74-109">If you don't have an Office Mac build, you can get one by joining the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

<span data-ttu-id="1cf74-110">Pour commencer, ouvrez un terminal, puis définissez la propriété `OfficeWebAddinDeveloperExtras` pour l’application Office pertinente comme suit :</span><span class="sxs-lookup"><span data-stu-id="1cf74-110">To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:</span></span>

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

<span data-ttu-id="1cf74-111">Ensuite, ouvrez l’application Office et[insérez votre complément](sideload-an-office-add-in-on-ipad-and-mac.md).</span><span class="sxs-lookup"><span data-stu-id="1cf74-111">Then, open the Office application and [sideload your add-in](sideload-an-office-add-in-on-ipad-and-mac.md).</span></span> <span data-ttu-id="1cf74-112">Cliquez sur le complément. Vous devriez voir l’option **Inspecter l’élément** s’afficher dans le menu contextuel.</span><span class="sxs-lookup"><span data-stu-id="1cf74-112">Right-click the add-in and you should see an **Inspect Element** option in the context menu.</span></span> <span data-ttu-id="1cf74-113">Sélectionnez cette option pour afficher l’inspecteur dans lequel vous pouvez définir des points d’arrêt et déboguer votre complément.</span><span class="sxs-lookup"><span data-stu-id="1cf74-113">Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="1cf74-114">Si vous essayez d’utiliser l’inspecteur et si la boîte de dialogue scintille, mettez Office à jour vers la dernière version.</span><span class="sxs-lookup"><span data-stu-id="1cf74-114">If you're trying to use the inspector and the dialog flickers, update Office to the latest version.</span></span> <span data-ttu-id="1cf74-115">Si cela ne résout pas le problème de scintillement, essayez la solution de contournement suivante :</span><span class="sxs-lookup"><span data-stu-id="1cf74-115">If that doesn't resolve the flickering, try the following workaround:</span></span>
> 1. <span data-ttu-id="1cf74-116">Pour réduire la taille de la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="1cf74-116">Reduce the size of the dialog.</span></span>
> 2. <span data-ttu-id="1cf74-117">Sélectionnez l’option **Inspecter l’élément** qui ouvre une nouvelle fenêtre.</span><span class="sxs-lookup"><span data-stu-id="1cf74-117">Choose **Inspect Element**, which opens in a new window.</span></span>
> 3. <span data-ttu-id="1cf74-118">Redimensionner la boîte de dialogue à sa taille d’origine.</span><span class="sxs-lookup"><span data-stu-id="1cf74-118">Resize the dialog to its original size.</span></span>
> 4. <span data-ttu-id="1cf74-119">Utiliser l’inspecteur comme requis.</span><span class="sxs-lookup"><span data-stu-id="1cf74-119">Use the inspector as required.</span></span>

## <a name="clearing-the-office-applications-cache-on-a-mac"></a><span data-ttu-id="1cf74-120">Effacement du cache de l’application Office sur un ordinateur Mac</span><span class="sxs-lookup"><span data-stu-id="1cf74-120">Clearing the Office application's cache on a Mac</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]
