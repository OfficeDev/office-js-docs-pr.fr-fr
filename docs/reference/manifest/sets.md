---
title: Élément Sets dans le fichier manifeste
description: L’élément sets spécifie l’ensemble minimal d’API JavaScript pour Office requis pour l’activation de votre complément Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: bd8f8311bb06a8e9e98fc408aece6395ab5643b1
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641423"
---
# <a name="sets-element"></a><span data-ttu-id="8d746-103">Élément Sets</span><span class="sxs-lookup"><span data-stu-id="8d746-103">Sets element</span></span>

<span data-ttu-id="8d746-104">Spécifie le sous-ensemble minimal de l’API JavaScript Office requise pour l’activation de votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="8d746-104">Specifies the minimum subset of the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="8d746-105">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="8d746-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="8d746-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="8d746-106">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="8d746-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="8d746-107">Contained in</span></span>

[<span data-ttu-id="8d746-108">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8d746-108">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="8d746-109">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="8d746-109">Can contain</span></span>

[<span data-ttu-id="8d746-110">Ensemble</span><span class="sxs-lookup"><span data-stu-id="8d746-110">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="8d746-111">Attributs</span><span class="sxs-lookup"><span data-stu-id="8d746-111">Attributes</span></span>

|<span data-ttu-id="8d746-112">Attribut</span><span class="sxs-lookup"><span data-stu-id="8d746-112">Attribute</span></span>|<span data-ttu-id="8d746-113">Type</span><span class="sxs-lookup"><span data-stu-id="8d746-113">Type</span></span>|<span data-ttu-id="8d746-114">Requis</span><span class="sxs-lookup"><span data-stu-id="8d746-114">Required</span></span>|<span data-ttu-id="8d746-115">Description</span><span class="sxs-lookup"><span data-stu-id="8d746-115">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="8d746-116">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="8d746-116">DefaultMinVersion</span></span>|<span data-ttu-id="8d746-117">chaîne</span><span class="sxs-lookup"><span data-stu-id="8d746-117">string</span></span>|<span data-ttu-id="8d746-118">facultatif</span><span class="sxs-lookup"><span data-stu-id="8d746-118">optional</span></span>|<span data-ttu-id="8d746-119">Spécifie la valeur par défaut de l’attribut **MinVersion** pour tous les éléments [Set](set.md) enfants.</span><span class="sxs-lookup"><span data-stu-id="8d746-119">Specifies the default **MinVersion** attribute value for all child [Set](set.md) elements.</span></span> <span data-ttu-id="8d746-120">La valeur par défaut est « 1.1 ».</span><span class="sxs-lookup"><span data-stu-id="8d746-120">The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="8d746-121">Remarques</span><span class="sxs-lookup"><span data-stu-id="8d746-121">Remarks</span></span>

<span data-ttu-id="8d746-122">Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="8d746-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="8d746-123">Pour plus d’informations sur l’attribut **MinVersion** de l’élément **Set** et sur l’attribut **DefaultMinVersion** de l’élément **sets** , voir [Set the requirements ELEMENT dans le manifeste](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="8d746-123">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span>

