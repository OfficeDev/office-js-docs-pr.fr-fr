---
title: Référence de l’API JavaScript pour OneNote
description: En savoir plus sur l’API JavaScript pour OneNote
ms.date: 07/28/2020
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: d917d71cd9d3f4fadbab91a434a177c45b54c6f2
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349111"
---
# <a name="onenote-javascript-api-overview"></a><span data-ttu-id="6f52a-103">Référence de l’API JavaScript pour OneNote</span><span class="sxs-lookup"><span data-stu-id="6f52a-103">OneNote JavaScript API overview</span></span>

<span data-ttu-id="6f52a-104">Un complément OneNote interagit avec des objets dans OneNote sur le web à l’aide de l’API JavaScript pour Office, qui inclut deux modèles objet JavaScript :</span><span class="sxs-lookup"><span data-stu-id="6f52a-104">A OneNote add-in interacts with objects in OneNote on the web by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="6f52a-p101">**API JavaScript OneNote** : il s’agit des [API spécifiques à l’application](../../develop/application-specific-api-model.md) pour OneNote. Introduite avec Office 2016, l’[API OneNote JavaScript](/javascript/api/onenote) fournit des objets fortement typés que vous pouvez utiliser pour accéder aux objets dans OneNote sur le web.</span><span class="sxs-lookup"><span data-stu-id="6f52a-p101">**OneNote JavaScript API**: These are the [application-specific APIs](../../develop/application-specific-api-model.md) for OneNote. Introduced with Office 2016, the [OneNote JavaScript API](/javascript/api/onenote) provides strongly-typed objects that you can use to access objects in OneNote on the web.</span></span>

* <span data-ttu-id="6f52a-107">**API communes** : incluses dans Office 2013, les [API communes](/javascript/api/office) peuvent être utilisées pour accéder à des fonctionnalités telles qu’une interface utilisateur, des boîtes de dialogue et des paramètres du client, qui sont communes à plusieurs types d’applications Office.</span><span class="sxs-lookup"><span data-stu-id="6f52a-107">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="6f52a-p102">Cette section de la documentation se concentre sur l’API JavaScript OneNote, que vous allez utiliser pour développer la majorité des fonctionnalités dans les compléments qui ciblent OneNote sur le web. Pour plus d’informations sur l’API commune, consultez [Modèle objet d’API JavaScript commun](../../develop/office-javascript-api-object-model.md).</span><span class="sxs-lookup"><span data-stu-id="6f52a-p102">This section of the documentation focuses on the OneNote JavaScript API, which you'll use to develop the majority of functionality in add-ins that target OneNote on the web. For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span>

## <a name="learn-programming-concepts"></a><span data-ttu-id="6f52a-110">Découvrir les concepts de programmation</span><span class="sxs-lookup"><span data-stu-id="6f52a-110">Learn programming concepts</span></span>

<span data-ttu-id="6f52a-111">Pour plus d’informations sur les concepts de programmation importants, consultez les articles suivants.</span><span class="sxs-lookup"><span data-stu-id="6f52a-111">See the following articles for information about important programming concepts.</span></span>

* [<span data-ttu-id="6f52a-112">Vue d’ensemble de la programmation de l’API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="6f52a-112">OneNote JavaScript API programming overview</span></span>](../../onenote/onenote-add-ins-programming-overview.md)
* [<span data-ttu-id="6f52a-113">Utiliser du contenu de page OneNote</span><span class="sxs-lookup"><span data-stu-id="6f52a-113">Work with OneNote page content</span></span>](../../onenote/onenote-add-ins-page-content.md)

## <a name="learn-about-api-capabilities"></a><span data-ttu-id="6f52a-114">En savoir plus sur les fonctionnalités des API</span><span class="sxs-lookup"><span data-stu-id="6f52a-114">Learn about API capabilities</span></span>

<span data-ttu-id="6f52a-115">Pour avoir une expérience pratique de l’utilisation de l’API JavaScript OneNote afin d’interagir avec du contenu dans OneNote sur le web, suivez le [Démarrage rapide du complément OneNote](../../quickstarts/onenote-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="6f52a-115">For hands-on experience using the OneNote JavaScript API to interact with content in OneNote on the web, complete the [OneNote add-in quick start](../../quickstarts/onenote-quickstart.md).</span></span>

<span data-ttu-id="6f52a-116">Pour en savoir plus sur le modèle d'objet de l’API JavaScript pour OneNote, consultez la [Documentation de référence de l’API JavaScript pour OneNote](/javascript/api/onenote).</span><span class="sxs-lookup"><span data-stu-id="6f52a-116">For detailed information about the OneNote JavaScript API object model, see the [OneNote JavaScript API reference documentation](/javascript/api/onenote).</span></span>

## <a name="see-also"></a><span data-ttu-id="6f52a-117">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="6f52a-117">See also</span></span>

* [<span data-ttu-id="6f52a-118">Documentation sur les compléments OneNote</span><span class="sxs-lookup"><span data-stu-id="6f52a-118">OneNote add-ins documentation</span></span>](../../onenote/index.yml)
* [<span data-ttu-id="6f52a-119">Présentation des compléments OneNote</span><span class="sxs-lookup"><span data-stu-id="6f52a-119">OneNote add-ins overview</span></span>](../../onenote/onenote-add-ins-programming-overview.md)
* [<span data-ttu-id="6f52a-120">Référence de l’API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="6f52a-120">OneNote JavaScript API reference</span></span>](/javascript/api/onenote)
* [<span data-ttu-id="6f52a-121">Application cliente Office et disponibilité de la plateforme pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="6f52a-121">Office client application and platform availability for Office Add-ins</span></span>](../../overview/office-add-in-availability.md)
