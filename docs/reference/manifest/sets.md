---
title: Élément Sets dans le fichier manifeste
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 13777e54ec6bd2d97fa35609ebe194ed85ffa1b8
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450421"
---
# <a name="sets-element"></a><span data-ttu-id="916af-102">Sets, élément</span><span class="sxs-lookup"><span data-stu-id="916af-102">Sets element</span></span>

<span data-ttu-id="916af-103">Spécifie le sous-ensemble minimal de l’API JavaScript pour Office nécessaire à l’activation de votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="916af-103">Specifies the minimum subset of the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="916af-104">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="916af-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="916af-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="916af-105">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="916af-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="916af-106">Contained in</span></span>

[<span data-ttu-id="916af-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="916af-107">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="916af-108">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="916af-108">Can contain</span></span>

[<span data-ttu-id="916af-109">Ensemble</span><span class="sxs-lookup"><span data-stu-id="916af-109">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="916af-110">Attributs</span><span class="sxs-lookup"><span data-stu-id="916af-110">Attributes</span></span>

|<span data-ttu-id="916af-111">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="916af-111">**Attribute**</span></span>|<span data-ttu-id="916af-112">**Type**</span><span class="sxs-lookup"><span data-stu-id="916af-112">**Type**</span></span>|<span data-ttu-id="916af-113">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="916af-113">**Required**</span></span>|<span data-ttu-id="916af-114">**Description**</span><span class="sxs-lookup"><span data-stu-id="916af-114">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="916af-115">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="916af-115">DefaultMinVersion</span></span>|<span data-ttu-id="916af-116">chaîne</span><span class="sxs-lookup"><span data-stu-id="916af-116">string</span></span>|<span data-ttu-id="916af-117">facultatif</span><span class="sxs-lookup"><span data-stu-id="916af-117">optional</span></span>|<span data-ttu-id="916af-p101">Spécifie la valeur de l’attribut **MinVersion** par défaut pour tous les éléments [Set](set.md) enfants. La valeur par défaut est « 1.1 ».</span><span class="sxs-lookup"><span data-stu-id="916af-p101">Specifies the default  **MinVersion** attribute value for all child [Set](set.md) elements. The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="916af-120">Remarques</span><span class="sxs-lookup"><span data-stu-id="916af-120">Remarks</span></span>

<span data-ttu-id="916af-121">Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="916af-121">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="916af-122">Pour plus d'informations sur l’attribut **MinVersion** de l’élément **Set** et sur l’attribut **DefaultMinVersion** de l’élément **Sets**, voir l’article relatif à la [définition de l’élément Requirements dans le manifeste](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="916af-122">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

