---
title: Ensemble de conditions requises de l’API du complément Outlook 1.4
description: Fonctionnalités et API introduites pour les Outlook et les API JavaScript Office dans le cadre de l’API de boîte aux lettres 1.4.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 19d77784926ac09d5620eb36242701da59b39f09
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591015"
---
# <a name="outlook-add-in-api-requirement-set-14"></a><span data-ttu-id="74af9-103">Ensemble de conditions requises de l’API du complément Outlook 1.4</span><span class="sxs-lookup"><span data-stu-id="74af9-103">Outlook add-in API requirement set 1.4</span></span>

<span data-ttu-id="74af9-104">Le sous-ensemble d’API de Outlook de l’API JavaScript Office inclut des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un Outlook.</span><span class="sxs-lookup"><span data-stu-id="74af9-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="74af9-105">Dans cette documentation, l’[ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) présenté est différent de l’ensemble de conditions requises de la version précédente.</span><span class="sxs-lookup"><span data-stu-id="74af9-105">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-14"></a><span data-ttu-id="74af9-106">Nouveautés de la version 1.4</span><span class="sxs-lookup"><span data-stu-id="74af9-106">What's new in 1.4?</span></span>

<span data-ttu-id="74af9-107">L’ensemble de conditions requises 1.4 inclut toutes les fonctionnalités de l’ensemble de conditions [requises 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md).</span><span class="sxs-lookup"><span data-stu-id="74af9-107">Requirement set 1.4 includes all of the features of [requirement set 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md).</span></span> <span data-ttu-id="74af9-108">Il comprend en plus l’accès à l’espace de noms `Office.ui`.</span><span class="sxs-lookup"><span data-stu-id="74af9-108">It added access to the `Office.ui` namespace.</span></span>

### <a name="change-log"></a><span data-ttu-id="74af9-109">Journal des modifications</span><span class="sxs-lookup"><span data-stu-id="74af9-109">Change log</span></span>

- <span data-ttu-id="74af9-110">Ajout [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): affiche une boîte de dialogue dans Office application.</span><span class="sxs-lookup"><span data-stu-id="74af9-110">Added [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): Displays a dialog box in an Office application.</span></span>
- <span data-ttu-id="74af9-111">Ajout de la méthode[Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-): Remet un message de la part de la boîte de dialogue à sa page parent/d’ouverture.</span><span class="sxs-lookup"><span data-stu-id="74af9-111">Added [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-): Delivers a message from the dialog box to its parent/opener page.</span></span>
- <span data-ttu-id="74af9-112">Ajout de l’objet [Dialog](/javascript/api/office/office.dialog): objet renvoyé lorsque la méthode [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-)est appelée.</span><span class="sxs-lookup"><span data-stu-id="74af9-112">Added [Dialog](/javascript/api/office/office.dialog) object: The object that is returned when the [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) method is called.</span></span>

## <a name="see-also"></a><span data-ttu-id="74af9-113">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="74af9-113">See also</span></span>

- [<span data-ttu-id="74af9-114">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="74af9-114">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="74af9-115">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="74af9-115">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="74af9-116">Prise en main</span><span class="sxs-lookup"><span data-stu-id="74af9-116">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="74af9-117">Ensembles de conditions requises et clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="74af9-117">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
