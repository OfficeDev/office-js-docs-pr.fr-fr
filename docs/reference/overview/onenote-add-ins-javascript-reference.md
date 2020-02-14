---
title: Référence de l’API JavaScript pour OneNote
description: ''
ms.date: 07/05/2019
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: 3ad992bd59c33d9d8b724893db49a6e623fd1ee3
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950977"
---
# <a name="onenote-javascript-api-overview"></a><span data-ttu-id="e3eb6-102">Référence de l’API JavaScript pour OneNote</span><span class="sxs-lookup"><span data-stu-id="e3eb6-102">OneNote JavaScript API overview</span></span>

<span data-ttu-id="e3eb6-103">Un complément OneNote interagit avec des objets dans OneNote sur le web à l’aide de l’interface API JavaScript pour Office, qui inclut deux modèles d’objets JavaScript :</span><span class="sxs-lookup"><span data-stu-id="e3eb6-103">A OneNote add-in interacts with objects in OneNote on the web by using the JavaScript API for Office, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="e3eb6-104">**API JavaScript pour OneNote** : inclut dans Office 2016, l’[API JavaScript pour OneNote](/javascript/api/onenote) fournit des objets fortement typés que vous pouvez utiliser pour accéder à des objets et à des métadonnées dans un document OneNote sur le web.</span><span class="sxs-lookup"><span data-stu-id="e3eb6-104">**OneNote JavaScript API**: Introduced with Office 2016, the [OneNote JavaScript API](/javascript/api/onenote) provides strongly-typed objects that you can use to access objects in OneNote on the web.</span></span> 

* <span data-ttu-id="e3eb6-105">**API communes** : incluses dans Office 2013, les [API communes](/javascript/api/office) peuvent être utilisées pour accéder à des fonctionnalités telles qu’une interface utilisateur, des boîtes de dialogue et des paramètres du client, qui sont communes à plusieurs types d’applications Office.</span><span class="sxs-lookup"><span data-stu-id="e3eb6-105">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="e3eb6-106">Cette section de la documentation traite de l’API JavaScript pour OneNote, que vous allez utiliser pour développer la majorité des fonctionnalités des compléments utilisés dans OneNote sur le web.</span><span class="sxs-lookup"><span data-stu-id="e3eb6-106">This section of the documentation focuses on the OneNote JavaScript API, which you'll use to develop the majority of functionality in add-ins that target OneNote on the web.</span></span> <span data-ttu-id="e3eb6-107">Pour plus d’informations sur les API comm.unes, voir le [Modèle d’objet API JavaScript pour Office](../../develop/office-javascript-api-object-model.md).</span><span class="sxs-lookup"><span data-stu-id="e3eb6-107">For information about the Common API, see [Office JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span> 

## <a name="learn-programming-concepts"></a><span data-ttu-id="e3eb6-108">Découvrir les concepts de programmation</span><span class="sxs-lookup"><span data-stu-id="e3eb6-108">Learn programming concepts</span></span>

<span data-ttu-id="e3eb6-109">Pour plus d’informations sur les concepts de programmation essentiels, consultez les articles suivants :</span><span class="sxs-lookup"><span data-stu-id="e3eb6-109">See the following articles for information about important programming concepts:</span></span>

- [<span data-ttu-id="e3eb6-110">Vue d’ensemble de la programmation de l’API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="e3eb6-110">OneNote JavaScript API programming overview</span></span>](../../onenote/onenote-add-ins-programming-overview.md)

- [<span data-ttu-id="e3eb6-111">Utiliser du contenu de page OneNote</span><span class="sxs-lookup"><span data-stu-id="e3eb6-111">Work with OneNote page content</span></span>](../../onenote/onenote-add-ins-page-content.md)

## <a name="learn-about-api-capabilities"></a><span data-ttu-id="e3eb6-112">En savoir plus sur les fonctionnalités des API</span><span class="sxs-lookup"><span data-stu-id="e3eb6-112">Learn about API capabilities</span></span>

<span data-ttu-id="e3eb6-113">Pour avoir une expérience pratique de l’utilisation de l’API JavaScript OneNote afin d’interagir avec du contenu dans OneNote sur le web, suivez le [Démarrage rapide du complément OneNote](../../quickstarts/onenote-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="e3eb6-113">For hands-on experience using the OneNote JavaScript API to interact with content in OneNote on the web, complete the [OneNote add-in quick start](../../quickstarts/onenote-quickstart.md).</span></span> 

<span data-ttu-id="e3eb6-114">Pour en savoir plus sur le modèle d'objet de l’API JavaScript pour OneNote, consultez la [Documentation de référence de l’API JavaScript pour OneNote](/javascript/api/onenote).</span><span class="sxs-lookup"><span data-stu-id="e3eb6-114">For detailed information about the OneNote JavaScript API object model, see the [OneNote JavaScript API reference documentation](/javascript/api/onenote).</span></span>

## <a name="see-also"></a><span data-ttu-id="e3eb6-115">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="e3eb6-115">See also</span></span>

- [<span data-ttu-id="e3eb6-116">Documentation sur les compléments OneNote</span><span class="sxs-lookup"><span data-stu-id="e3eb6-116">OneNote add-ins documentation</span></span>](../../onenote/index.md)
- [<span data-ttu-id="e3eb6-117">Présentation des compléments OneNote</span><span class="sxs-lookup"><span data-stu-id="e3eb6-117">OneNote add-ins overview</span></span>](../../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="e3eb6-118">Référence de l’API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="e3eb6-118">OneNote JavaScript API reference</span></span>](/javascript/api/onenote)
- [<span data-ttu-id="e3eb6-119">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="e3eb6-119">Office Add-in host and platform availability</span></span>](../../overview/office-add-in-availability.md)

