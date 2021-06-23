---
title: Boîtes de dialogue dans les compléments Office
description: Découvrez les meilleures pratiques pour la conception visuelle des boîtes de dialogue dans Office des applications.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d674b747effa57b8a75b79f98f5ff78ccc8a92a4
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076335"
---
# <a name="dialog-boxes-in-office-add-ins"></a><span data-ttu-id="ed2e3-103">Boîtes de dialogue dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="ed2e3-103">Dialog boxes in Office Add-ins</span></span>

<span data-ttu-id="ed2e3-p101">Les boîtes de dialogue sont des surfaces qui flottent au-dessus de la fenêtre active de l’application Office. Vous pouvez utiliser les boîtes de dialogue afin de fournir un espace supplémentaire sur l’écran pour les tâches comme les pages de connexion impossibles à ouvrir directement dans un volet des tâches, ou pour les demandes de confirmation d’une action effectuée par un utilisateur, ou pour afficher des vidéos qui peuvent être trop petites si confinées à un volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="ed2e3-p101">Dialog boxes are surfaces that float above the active Office application window. You can use dialog boxes to provide additional screen space for tasks such as sign-in pages that can't be opened directly in a task pane or requests to confirm an action taken by a user, or to show videos that might be too small if confined to a task pane.</span></span>

<span data-ttu-id="ed2e3-106">*Figure 1. Mise en page type pour une boîte de dialogue*</span><span class="sxs-lookup"><span data-stu-id="ed2e3-106">*Figure 1. Typical layout for a dialog box*</span></span>

![Disposition type d’une boîte de dialogue affichée dans une application Office’application.](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a><span data-ttu-id="ed2e3-108">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="ed2e3-108">Best practices</span></span>

|<span data-ttu-id="ed2e3-109">À faire</span><span class="sxs-lookup"><span data-stu-id="ed2e3-109">Do</span></span>|<span data-ttu-id="ed2e3-110">À ne pas faire</span><span class="sxs-lookup"><span data-stu-id="ed2e3-110">Don't</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="ed2e3-111">Inclure un titre descriptif qui inclut le nom de votre complément, ainsi que la tâche en cours.</span><span class="sxs-lookup"><span data-stu-id="ed2e3-111">Include a descriptive title that includes your add-in name along with the current task.</span></span></li></ul>|<ul><li><span data-ttu-id="ed2e3-112">Ne pas ajouter le nom de votre société au titre.</span><span class="sxs-lookup"><span data-stu-id="ed2e3-112">Don't append your company name to the title.</span></span></li></ul>|
||<ul><li><span data-ttu-id="ed2e3-113">Ne pas ouvrir une boîte de dialogue, sauf si le scénario l’exige.</span><span class="sxs-lookup"><span data-stu-id="ed2e3-113">Don't open a dialog box unless the scenario requires it.</span></span></li></ul>|

## <a name="implementation"></a><span data-ttu-id="ed2e3-114">Implémentation</span><span class="sxs-lookup"><span data-stu-id="ed2e3-114">Implementation</span></span>

<span data-ttu-id="ed2e3-115">Pour voir un exemple relatif à l’implémentation d’une boîte de dialogue, consultez [Exemple d’API de boîte de dialogue de complément Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) dans GitHub.</span><span class="sxs-lookup"><span data-stu-id="ed2e3-115">For a sample that implements a dialog box, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) in GitHub.</span></span>

## <a name="see-also"></a><span data-ttu-id="ed2e3-116">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ed2e3-116">See also</span></span>

- [<span data-ttu-id="ed2e3-117">Dialog object</span><span class="sxs-lookup"><span data-stu-id="ed2e3-117">Dialog object</span></span>](/javascript/api/office/office.dialog)
- [<span data-ttu-id="ed2e3-118">Modèles de conception de l’expérience utilisateur pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="ed2e3-118">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)
