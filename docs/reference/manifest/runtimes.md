---
title: Runtimes dans le fichier manifeste (aperçu)
description: ''
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 17e53b53d55ea9547cdfc5c4f89f8f4c3a7ab75e
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/26/2020
ms.locfileid: "42283872"
---
# <a name="runtimes-element-preview"></a><span data-ttu-id="dd32c-102">Runtimes, élément (aperçu)</span><span class="sxs-lookup"><span data-stu-id="dd32c-102">Runtimes element (preview)</span></span>

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="dd32c-103">Spécifie le runtime de votre complément et active des fonctions personnalisées, des boutons du ruban et le volet des tâches pour utiliser le même Runtime JavaScript.</span><span class="sxs-lookup"><span data-stu-id="dd32c-103">Specifies the runtime of your add-in and enables custom functions, ribbon buttons, and the task pane to use the same JavaScript runtime.</span></span> <span data-ttu-id="dd32c-104">Enfant de l' `<Host>` élément dans votre fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="dd32c-104">Child of the `<Host>` element in your manifest file.</span></span> <span data-ttu-id="dd32c-105">Pour plus d’informations, reportez-vous [à la rubrique Configure Your Excel Add-in to use a Shared JavaScript Runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="dd32c-105">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="dd32c-106">**Type de complément :** volet Office</span><span class="sxs-lookup"><span data-stu-id="dd32c-106">**Add-in type:** Task pane</span></span>

> [!IMPORTANT]
> <span data-ttu-id="dd32c-107">Le runtime partagé est actuellement en préversion et n’est disponible que sur Excel sur Windows.</span><span class="sxs-lookup"><span data-stu-id="dd32c-107">Shared runtime is currently in preview and are only available on Excel on Windows.</span></span> <span data-ttu-id="dd32c-108">Pour essayer les fonctionnalités d’aperçu, vous devrez rejoindre [Office Insider](https://insider.office.com/).</span><span class="sxs-lookup"><span data-stu-id="dd32c-108">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>

## <a name="syntax"></a><span data-ttu-id="dd32c-109">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="dd32c-109">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="dd32c-110">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="dd32c-110">Contained in</span></span> 
[<span data-ttu-id="dd32c-111">Host</span><span class="sxs-lookup"><span data-stu-id="dd32c-111">Host</span></span>](./host.md)

## <a name="child-elements"></a><span data-ttu-id="dd32c-112">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="dd32c-112">Child elements</span></span>

|  <span data-ttu-id="dd32c-113">Élément</span><span class="sxs-lookup"><span data-stu-id="dd32c-113">Element</span></span> |  <span data-ttu-id="dd32c-114">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="dd32c-114">Required</span></span>  |  <span data-ttu-id="dd32c-115">Description</span><span class="sxs-lookup"><span data-stu-id="dd32c-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="dd32c-116">**Runtime**</span><span class="sxs-lookup"><span data-stu-id="dd32c-116">**Runtime**</span></span>     | <span data-ttu-id="dd32c-117">Oui</span><span class="sxs-lookup"><span data-stu-id="dd32c-117">Yes</span></span> |  <span data-ttu-id="dd32c-118">Le runtime de votre complément.</span><span class="sxs-lookup"><span data-stu-id="dd32c-118">The runtime for your add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="dd32c-119">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="dd32c-119">See also</span></span>

- [<span data-ttu-id="dd32c-120">Runtime</span><span class="sxs-lookup"><span data-stu-id="dd32c-120">Runtime</span></span>](runtime.md)
