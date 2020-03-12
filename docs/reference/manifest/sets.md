---
title: Élément Sets dans le fichier manifeste
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 80f8a74b64186496ac1579b283b3e2976978328b
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596486"
---
# <a name="sets-element"></a><span data-ttu-id="3ad49-102">Sets, élément</span><span class="sxs-lookup"><span data-stu-id="3ad49-102">Sets element</span></span>

<span data-ttu-id="3ad49-103">Spécifie le sous-ensemble minimal de l’API JavaScript Office requise pour l’activation de votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="3ad49-103">Specifies the minimum subset of the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="3ad49-104">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="3ad49-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="3ad49-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="3ad49-105">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="3ad49-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="3ad49-106">Contained in</span></span>

[<span data-ttu-id="3ad49-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="3ad49-107">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="3ad49-108">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="3ad49-108">Can contain</span></span>

[<span data-ttu-id="3ad49-109">Ensemble</span><span class="sxs-lookup"><span data-stu-id="3ad49-109">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="3ad49-110">Attributs</span><span class="sxs-lookup"><span data-stu-id="3ad49-110">Attributes</span></span>

|<span data-ttu-id="3ad49-111">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="3ad49-111">**Attribute**</span></span>|<span data-ttu-id="3ad49-112">**Type**</span><span class="sxs-lookup"><span data-stu-id="3ad49-112">**Type**</span></span>|<span data-ttu-id="3ad49-113">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="3ad49-113">**Required**</span></span>|<span data-ttu-id="3ad49-114">**Description**</span><span class="sxs-lookup"><span data-stu-id="3ad49-114">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="3ad49-115">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="3ad49-115">DefaultMinVersion</span></span>|<span data-ttu-id="3ad49-116">chaîne</span><span class="sxs-lookup"><span data-stu-id="3ad49-116">string</span></span>|<span data-ttu-id="3ad49-117">facultatif</span><span class="sxs-lookup"><span data-stu-id="3ad49-117">optional</span></span>|<span data-ttu-id="3ad49-118">Spécifie la valeur par défaut de l’attribut **MinVersion** pour tous les éléments [Set](set.md) enfants.</span><span class="sxs-lookup"><span data-stu-id="3ad49-118">Specifies the default **MinVersion** attribute value for all child [Set](set.md) elements.</span></span> <span data-ttu-id="3ad49-119">La valeur par défaut est « 1.1 ».</span><span class="sxs-lookup"><span data-stu-id="3ad49-119">The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="3ad49-120">Remarques</span><span class="sxs-lookup"><span data-stu-id="3ad49-120">Remarks</span></span>

<span data-ttu-id="3ad49-121">Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="3ad49-121">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="3ad49-122">Pour plus d’informations sur l’attribut **MinVersion** de l’élément **Set** et sur l’attribut **DefaultMinVersion** de l’élément **sets** , voir [Set the requirements ELEMENT dans le manifeste](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="3ad49-122">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span>

