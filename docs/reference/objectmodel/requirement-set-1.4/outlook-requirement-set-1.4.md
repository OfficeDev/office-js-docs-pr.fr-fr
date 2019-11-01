---
title: Ensemble de conditions requises de l’API du complément Outlook 1.4
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: ded3aa8cdcde1132f074f55356e64bb4468919ff
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/31/2019
ms.locfileid: "37901940"
---
# <a name="outlook-add-in-api-requirement-set-14"></a><span data-ttu-id="745fa-102">Ensemble de conditions requises de l’API du complément Outlook 1.4</span><span class="sxs-lookup"><span data-stu-id="745fa-102">Outlook add-in API requirement set 1.4</span></span>

<span data-ttu-id="745fa-103">Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="745fa-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="745fa-104">Dans cette documentation, l’[ensemble de conditions requises](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) présenté est différent de l’ensemble de conditions requises de la version précédente.</span><span class="sxs-lookup"><span data-stu-id="745fa-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-14"></a><span data-ttu-id="745fa-105">Nouveautés de la version 1.4</span><span class="sxs-lookup"><span data-stu-id="745fa-105">What's new in 1.4?</span></span>

<span data-ttu-id="745fa-p101">L’ensemble de conditions requises de la version 1.4 comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). Il comprend en plus l’accès à l’espace de noms `Office.ui`.</span><span class="sxs-lookup"><span data-stu-id="745fa-p101">Requirement set 1.4 includes all of the features of [Requirement set 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). It added access to the `Office.ui` namespace.</span></span>

### <a name="change-log"></a><span data-ttu-id="745fa-108">Journal des modifications</span><span class="sxs-lookup"><span data-stu-id="745fa-108">Change log</span></span>

- <span data-ttu-id="745fa-109">Ajout de la méthode [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) : Affiche une boîte de dialogue dans un hôte Office.</span><span class="sxs-lookup"><span data-stu-id="745fa-109">Added [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): Displays a dialog box in an Office host.</span></span>
- <span data-ttu-id="745fa-110">Ajout de la méthode[Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-): Remet un message de la part de la boîte de dialogue à sa page parent/d’ouverture.</span><span class="sxs-lookup"><span data-stu-id="745fa-110">Added [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-): Delivers a message from the dialog box to its parent/opener page.</span></span>
- <span data-ttu-id="745fa-111">Ajout de l’objet [Dialog](/javascript/api/office/office.dialog): objet renvoyé lorsque la méthode [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-)est appelée.</span><span class="sxs-lookup"><span data-stu-id="745fa-111">Added [Dialog](/javascript/api/office/office.dialog) object: The object that is returned when the [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) method is called.</span></span>

## <a name="see-also"></a><span data-ttu-id="745fa-112">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="745fa-112">See also</span></span>

- [<span data-ttu-id="745fa-113">Compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="745fa-113">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="745fa-114">Exemples de code pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="745fa-114">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="745fa-115">Prise en main</span><span class="sxs-lookup"><span data-stu-id="745fa-115">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="745fa-116">Ensembles de conditions requises et clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="745fa-116">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
