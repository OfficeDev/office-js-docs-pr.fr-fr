---
title: Élément Sets dans le fichier manifeste
description: L’élément sets spécifie l’ensemble minimal d’API JavaScript pour Office requis pour l’activation de votre complément Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 8c1c97bfc2934ecf3cc20b472b29a03805603729
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608732"
---
# <a name="sets-element"></a><span data-ttu-id="84e82-103">Élément Sets</span><span class="sxs-lookup"><span data-stu-id="84e82-103">Sets element</span></span>

<span data-ttu-id="84e82-104">Spécifie le sous-ensemble minimal de l’API JavaScript Office requise pour l’activation de votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="84e82-104">Specifies the minimum subset of the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="84e82-105">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="84e82-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="84e82-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="84e82-106">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="84e82-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="84e82-107">Contained in</span></span>

[<span data-ttu-id="84e82-108">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="84e82-108">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="84e82-109">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="84e82-109">Can contain</span></span>

[<span data-ttu-id="84e82-110">Ensemble</span><span class="sxs-lookup"><span data-stu-id="84e82-110">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="84e82-111">Attributs</span><span class="sxs-lookup"><span data-stu-id="84e82-111">Attributes</span></span>

|<span data-ttu-id="84e82-112">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="84e82-112">**Attribute**</span></span>|<span data-ttu-id="84e82-113">**Type**</span><span class="sxs-lookup"><span data-stu-id="84e82-113">**Type**</span></span>|<span data-ttu-id="84e82-114">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="84e82-114">**Required**</span></span>|<span data-ttu-id="84e82-115">**Description**</span><span class="sxs-lookup"><span data-stu-id="84e82-115">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="84e82-116">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="84e82-116">DefaultMinVersion</span></span>|<span data-ttu-id="84e82-117">chaîne</span><span class="sxs-lookup"><span data-stu-id="84e82-117">string</span></span>|<span data-ttu-id="84e82-118">facultatif</span><span class="sxs-lookup"><span data-stu-id="84e82-118">optional</span></span>|<span data-ttu-id="84e82-119">Spécifie la valeur par défaut de l’attribut **MinVersion** pour tous les éléments [Set](set.md) enfants.</span><span class="sxs-lookup"><span data-stu-id="84e82-119">Specifies the default **MinVersion** attribute value for all child [Set](set.md) elements.</span></span> <span data-ttu-id="84e82-120">La valeur par défaut est « 1.1 ».</span><span class="sxs-lookup"><span data-stu-id="84e82-120">The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="84e82-121">Remarques</span><span class="sxs-lookup"><span data-stu-id="84e82-121">Remarks</span></span>

<span data-ttu-id="84e82-122">Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="84e82-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="84e82-123">Pour plus d’informations sur l’attribut **MinVersion** de l’élément **Set** et sur l’attribut **DefaultMinVersion** de l’élément **sets** , voir [Set the requirements ELEMENT dans le manifeste](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="84e82-123">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span>

