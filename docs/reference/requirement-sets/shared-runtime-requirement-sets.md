---
title: Ensembles de conditions requises pour l’exécution partagée
description: Spécifie les plateformes et les hôtes Office qui prennent en charge les API SharedRuntime.
ms.date: 02/11/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: dbb9d908154da074eaff6901c778adea168504a9
ms.sourcegitcommit: 7464eac3b54a6a6b65e27549a3ad603af6ee1011
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42315879"
---
# <a name="shared-runtime-requirement-sets"></a><span data-ttu-id="a149f-103">Ensembles de conditions requises pour l’exécution partagée</span><span class="sxs-lookup"><span data-stu-id="a149f-103">Shared runtime requirement sets</span></span>

<span data-ttu-id="a149f-p101">Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="a149f-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="a149f-107">Les parties d’un complément Office qui exécutent du code JavaScript, telles que des volets de tâches, des fichiers de fonctions lancés à partir de commandes de complément et des fonctions personnalisées Excel, peuvent partager un seul Runtime JavaScript.</span><span class="sxs-lookup"><span data-stu-id="a149f-107">Parts of an Office Add-in that run JavaScript code, such as task panes, function files launched from add-in commands, and Excel custom functions, can share a single JavaScript runtime.</span></span> <span data-ttu-id="a149f-108">Cela permet à toutes les parties de partager un ensemble de variables globales, de partager un ensemble de bibliothèques chargées et de communiquer les uns avec les autres sans avoir à transmettre de messages via un stockage persistant.</span><span class="sxs-lookup"><span data-stu-id="a149f-108">This enables all the parts to share a set of global variables, to share a set of loaded libraries, and to communicate with each other without having to pass messages through a persisted storage.</span></span>

<span data-ttu-id="a149f-109">Le tableau suivant répertorie l’ensemble de conditions requises SharedRuntime 1,1, les applications hôtes Office qui prennent en charge cet ensemble de conditions requises, ainsi que les numéros de build ou de version de l’application Office.</span><span class="sxs-lookup"><span data-stu-id="a149f-109">The following table lists the SharedRuntime 1.1 requirement set, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="a149f-110">Ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="a149f-110">Requirement set</span></span>  |  <span data-ttu-id="a149f-111">Office 2013 (ou version ultérieure) sur Windows</span><span class="sxs-lookup"><span data-stu-id="a149f-111">Office 2013 (or later) on Windows</span></span><br><span data-ttu-id="a149f-112">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="a149f-112">(one-time purchase)</span></span> | <span data-ttu-id="a149f-113">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="a149f-113">Office on Windows</span></span><br><span data-ttu-id="a149f-114">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="a149f-114">(connected to Office 365 subscription)</span></span>   |  <span data-ttu-id="a149f-115">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="a149f-115">Office on iPad</span></span><br><span data-ttu-id="a149f-116">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="a149f-116">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="a149f-117">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="a149f-117">Office on Mac</span></span><br><span data-ttu-id="a149f-118">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="a149f-118">(connected to Office 365 subscription)</span></span>  | <span data-ttu-id="a149f-119">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="a149f-119">Office on the web</span></span>  | <span data-ttu-id="a149f-120">Office Online Server</span><span class="sxs-lookup"><span data-stu-id="a149f-120">Office Online Server</span></span> |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="a149f-121">SharedRuntime 1,1</span><span class="sxs-lookup"><span data-stu-id="a149f-121">SharedRuntime 1.1</span></span>  | <span data-ttu-id="a149f-122">S/O</span><span class="sxs-lookup"><span data-stu-id="a149f-122">N/A</span></span> | <span data-ttu-id="a149f-123">Version 2002 (Build 12527,20092) ou version ultérieure</span><span class="sxs-lookup"><span data-stu-id="a149f-123">Version 2002 (Build 12527.20092) or later</span></span> | <span data-ttu-id="a149f-124">S/O</span><span class="sxs-lookup"><span data-stu-id="a149f-124">N/A</span></span> | <span data-ttu-id="a149f-125">16,35 ou version ultérieure</span><span class="sxs-lookup"><span data-stu-id="a149f-125">16.35 or later</span></span> | <span data-ttu-id="a149f-126">Février 2020</span><span class="sxs-lookup"><span data-stu-id="a149f-126">February 2020</span></span> | <span data-ttu-id="a149f-127">S/O</span><span class="sxs-lookup"><span data-stu-id="a149f-127">N/A</span></span> |

<span data-ttu-id="a149f-128">Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :</span><span class="sxs-lookup"><span data-stu-id="a149f-128">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

- [<span data-ttu-id="a149f-129">Numéros de version et de build des canaux de réception des mises à jour pour les clients Office 365</span><span class="sxs-lookup"><span data-stu-id="a149f-129">Version and build numbers of update channel releases for Office 365 clients</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="a149f-130">Quelle est la version d’Office que j’utilise ?</span><span class="sxs-lookup"><span data-stu-id="a149f-130">What version of Office am I using?</span></span>](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [<span data-ttu-id="a149f-131">Où trouver le numéro de version et de build pour une application cliente Office 365</span><span class="sxs-lookup"><span data-stu-id="a149f-131">Where you can find the version and build number for an Office 365 client application</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="a149f-132">Présentation d’Office Online Server</span><span class="sxs-lookup"><span data-stu-id="a149f-132">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="a149f-133">Ensembles de conditions requises des API communes pour Office</span><span class="sxs-lookup"><span data-stu-id="a149f-133">Office Common API requirement sets</span></span>

<span data-ttu-id="a149f-134">Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="a149f-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="a149f-135">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="a149f-135">See also</span></span>

- [<span data-ttu-id="a149f-136">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="a149f-136">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="a149f-137">Spécification des exigences en matière d’hôtes Office et d’API</span><span class="sxs-lookup"><span data-stu-id="a149f-137">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="a149f-138">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="a149f-138">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)