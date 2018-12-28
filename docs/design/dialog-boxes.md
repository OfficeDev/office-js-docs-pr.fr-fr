---
title: Boîtes de dialogue dans les compléments Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 5874e93c79d019adbd7ca5827a5b5e22f742190a
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457669"
---
# <a name="dialog-boxes-in-office-add-ins"></a><span data-ttu-id="3339f-102">Boîtes de dialogue dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="3339f-102">Dialog boxes in Office Add-ins</span></span>
 
<span data-ttu-id="3339f-p101">Les boîtes de dialogue sont des surfaces qui flottent au-dessus de la fenêtre active de l’application Office. Vous pouvez utiliser les boîtes de dialogue afin de fournir un espace supplémentaire sur l’écran pour les tâches comme les pages de connexion impossibles à ouvrir directement dans un volet des tâches, ou pour les demandes de confirmation d’une action effectuée par un utilisateur, ou pour afficher des vidéos qui peuvent être trop petites si confinées à un volet des tâches.</span><span class="sxs-lookup"><span data-stu-id="3339f-p101">Dialog boxes are surfaces that float above the active Office application window. You can use dialog boxes to provide additional screen space for tasks such as sign-in pages that can't be opened directly in a task pane or requests to confirm an action taken by a user, or to show videos that might be too small if confined to a task pane.</span></span>

<span data-ttu-id="3339f-105">*Figure 1. Mise en page type pour une boîte de dialogue*</span><span class="sxs-lookup"><span data-stu-id="3339f-105">*Figure 1. Typical layout for a dialog box*</span></span>

![Exemple d’image affichant une mise en page par défaut pour une boîte de dialogue](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a><span data-ttu-id="3339f-107">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="3339f-107">Best practices</span></span>

|<span data-ttu-id="3339f-108">**À faire**</span><span class="sxs-lookup"><span data-stu-id="3339f-108">**Do**</span></span>|<span data-ttu-id="3339f-109">**À ne pas faire**</span><span class="sxs-lookup"><span data-stu-id="3339f-109">**Don't**</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="3339f-110">Inclure un titre descriptif qui inclut le nom de votre complément, ainsi que la tâche en cours.</span><span class="sxs-lookup"><span data-stu-id="3339f-110">Include a descriptive title that includes your add-in name along with the current task.</span></span></li></ul>|<ul><li><span data-ttu-id="3339f-111">Ne pas ajouter le nom de votre société au titre.</span><span class="sxs-lookup"><span data-stu-id="3339f-111">Don't append your company name to the title.</span></span></li></ul>|
||<ul><li><span data-ttu-id="3339f-112">Ne pas ouvrir une boîte de dialogue, sauf si le scénario l’exige.</span><span class="sxs-lookup"><span data-stu-id="3339f-112">Don't open a dialog box unless the scenario requires it.</span></span></li></ul>|

## <a name="implementation"></a><span data-ttu-id="3339f-113">Implémentation</span><span class="sxs-lookup"><span data-stu-id="3339f-113">Implementation</span></span>

<span data-ttu-id="3339f-114">Pour voir un exemple relatif à l’implémentation d’une boîte de dialogue, consultez [Exemple d’API de boîte de dialogue de complément Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) dans GitHub.</span><span class="sxs-lookup"><span data-stu-id="3339f-114">For a sample that implements a dialog box, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) in GitHub.</span></span>

## <a name="see-also"></a><span data-ttu-id="3339f-115">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="3339f-115">See also</span></span>

- [<span data-ttu-id="3339f-116">Ressources de développement GitHub</span><span class="sxs-lookup"><span data-stu-id="3339f-116">GitHub Development Resources</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [<span data-ttu-id="3339f-117">Objet Dialogue</span><span class="sxs-lookup"><span data-stu-id="3339f-117">Dialog object</span></span>](https://docs.microsoft.com/javascript/api/office/office.dialog)


