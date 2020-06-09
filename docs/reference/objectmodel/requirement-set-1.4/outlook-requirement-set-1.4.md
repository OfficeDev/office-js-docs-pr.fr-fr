---
title: Ensemble de conditions requises de l’API du complément Outlook 1.4
description: Les fonctionnalités et les API qui ont été introduites pour les compléments Outlook et les API JavaScript Office dans le cadre de l’API de boîte aux lettres 1,4.
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: 6154acc357dd0e0e663d658c8de2d54b641e080a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609817"
---
# <a name="outlook-add-in-api-requirement-set-14"></a><span data-ttu-id="2bb17-103">Ensemble de conditions requises de l’API du complément Outlook 1.4</span><span class="sxs-lookup"><span data-stu-id="2bb17-103">Outlook add-in API requirement set 1.4</span></span>

<span data-ttu-id="2bb17-104">Le sous-ensemble d’API de complément Outlook de l’API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="2bb17-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="2bb17-105">Dans cette documentation, l’[ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) présenté est différent de l’ensemble de conditions requises de la version précédente.</span><span class="sxs-lookup"><span data-stu-id="2bb17-105">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-14"></a><span data-ttu-id="2bb17-106">Nouveautés de la version 1.4</span><span class="sxs-lookup"><span data-stu-id="2bb17-106">What's new in 1.4?</span></span>

<span data-ttu-id="2bb17-p101">L’ensemble de conditions requises de la version 1.4 comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). Il comprend en plus l’accès à l’espace de noms `Office.ui`.</span><span class="sxs-lookup"><span data-stu-id="2bb17-p101">Requirement set 1.4 includes all of the features of [Requirement set 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). It added access to the `Office.ui` namespace.</span></span>

### <a name="change-log"></a><span data-ttu-id="2bb17-109">Journal des modifications</span><span class="sxs-lookup"><span data-stu-id="2bb17-109">Change log</span></span>

- <span data-ttu-id="2bb17-110">Ajout de la méthode [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) : Affiche une boîte de dialogue dans un hôte Office.</span><span class="sxs-lookup"><span data-stu-id="2bb17-110">Added [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): Displays a dialog box in an Office host.</span></span>
- <span data-ttu-id="2bb17-111">Ajout de la méthode[Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-): Remet un message de la part de la boîte de dialogue à sa page parent/d’ouverture.</span><span class="sxs-lookup"><span data-stu-id="2bb17-111">Added [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-): Delivers a message from the dialog box to its parent/opener page.</span></span>
- <span data-ttu-id="2bb17-112">Ajout de l’objet [Dialog](/javascript/api/office/office.dialog): objet renvoyé lorsque la méthode [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-)est appelée.</span><span class="sxs-lookup"><span data-stu-id="2bb17-112">Added [Dialog](/javascript/api/office/office.dialog) object: The object that is returned when the [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) method is called.</span></span>

## <a name="see-also"></a><span data-ttu-id="2bb17-113">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="2bb17-113">See also</span></span>

- [<span data-ttu-id="2bb17-114">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="2bb17-114">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="2bb17-115">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="2bb17-115">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="2bb17-116">Prise en main</span><span class="sxs-lookup"><span data-stu-id="2bb17-116">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="2bb17-117">Ensembles de conditions requises et clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="2bb17-117">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
