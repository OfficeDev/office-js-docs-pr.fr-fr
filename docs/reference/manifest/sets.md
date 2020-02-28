---
title: Élément Sets dans le fichier manifeste
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 768f674b4afbd65df88825e871005f182d06f6ce
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325240"
---
# <a name="sets-element"></a><span data-ttu-id="8055b-102">Sets, élément</span><span class="sxs-lookup"><span data-stu-id="8055b-102">Sets element</span></span>

<span data-ttu-id="8055b-103">Spécifie le sous-ensemble minimal de l’API JavaScript Office requise pour l’activation de votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="8055b-103">Specifies the minimum subset of the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="8055b-104">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="8055b-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="8055b-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="8055b-105">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="8055b-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="8055b-106">Contained in</span></span>

[<span data-ttu-id="8055b-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8055b-107">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="8055b-108">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="8055b-108">Can contain</span></span>

[<span data-ttu-id="8055b-109">Ensemble</span><span class="sxs-lookup"><span data-stu-id="8055b-109">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="8055b-110">Attributs</span><span class="sxs-lookup"><span data-stu-id="8055b-110">Attributes</span></span>

|<span data-ttu-id="8055b-111">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="8055b-111">**Attribute**</span></span>|<span data-ttu-id="8055b-112">**Type**</span><span class="sxs-lookup"><span data-stu-id="8055b-112">**Type**</span></span>|<span data-ttu-id="8055b-113">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="8055b-113">**Required**</span></span>|<span data-ttu-id="8055b-114">**Description**</span><span class="sxs-lookup"><span data-stu-id="8055b-114">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="8055b-115">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="8055b-115">DefaultMinVersion</span></span>|<span data-ttu-id="8055b-116">chaîne</span><span class="sxs-lookup"><span data-stu-id="8055b-116">string</span></span>|<span data-ttu-id="8055b-117">facultatif</span><span class="sxs-lookup"><span data-stu-id="8055b-117">optional</span></span>|<span data-ttu-id="8055b-118">Spécifie la valeur par défaut de l’attribut **MinVersion** pour tous les éléments [Set](set.md) enfants.</span><span class="sxs-lookup"><span data-stu-id="8055b-118">Specifies the default **MinVersion** attribute value for all child [Set](set.md) elements.</span></span> <span data-ttu-id="8055b-119">La valeur par défaut est « 1.1 ».</span><span class="sxs-lookup"><span data-stu-id="8055b-119">The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="8055b-120">Remarques</span><span class="sxs-lookup"><span data-stu-id="8055b-120">Remarks</span></span>

<span data-ttu-id="8055b-121">Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="8055b-121">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="8055b-122">Pour plus d’informations sur l’attribut **MinVersion** de l’élément **Set** et sur l’attribut **DefaultMinVersion** de l’élément **sets** , voir [Set the requirements ELEMENT dans le manifeste](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="8055b-122">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

