---
title: Élément Set dans le fichier manifeste
description: L’élément Set spécifie un ensemble de conditions requises de l’API JavaScript pour Office requis pour l’activation de votre complément Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 608830e1ebc0d2e2d4c170b48bba00b3a19e87af
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641416"
---
# <a name="set-element"></a><span data-ttu-id="c3298-103">Élément Set</span><span class="sxs-lookup"><span data-stu-id="c3298-103">Set element</span></span>

<span data-ttu-id="c3298-104">Spécifie un ensemble de conditions requises de l’API JavaScript Office que votre complément Office doit activer.</span><span class="sxs-lookup"><span data-stu-id="c3298-104">Specifies a requirement set from the Office JavaScript API that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="c3298-105">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="c3298-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="c3298-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="c3298-106">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="c3298-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="c3298-107">Contained in</span></span>

[<span data-ttu-id="c3298-108">Ensembles</span><span class="sxs-lookup"><span data-stu-id="c3298-108">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="c3298-109">Attributs</span><span class="sxs-lookup"><span data-stu-id="c3298-109">Attributes</span></span>

|<span data-ttu-id="c3298-110">Attribut</span><span class="sxs-lookup"><span data-stu-id="c3298-110">Attribute</span></span>|<span data-ttu-id="c3298-111">Type</span><span class="sxs-lookup"><span data-stu-id="c3298-111">Type</span></span>|<span data-ttu-id="c3298-112">Requis</span><span class="sxs-lookup"><span data-stu-id="c3298-112">Required</span></span>|<span data-ttu-id="c3298-113">Description</span><span class="sxs-lookup"><span data-stu-id="c3298-113">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="c3298-114">Nom</span><span class="sxs-lookup"><span data-stu-id="c3298-114">Name</span></span>|<span data-ttu-id="c3298-115">string</span><span class="sxs-lookup"><span data-stu-id="c3298-115">string</span></span>|<span data-ttu-id="c3298-116">obligatoire</span><span class="sxs-lookup"><span data-stu-id="c3298-116">required</span></span>|<span data-ttu-id="c3298-117">Nom d’un [ensemble de conditions requises](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="c3298-117">The name of a [requirement set](../../develop/office-versions-and-requirement-sets.md).</span></span>|
|<span data-ttu-id="c3298-118">MinVersion</span><span class="sxs-lookup"><span data-stu-id="c3298-118">MinVersion</span></span>|<span data-ttu-id="c3298-119">chaîne</span><span class="sxs-lookup"><span data-stu-id="c3298-119">string</span></span>|<span data-ttu-id="c3298-120">facultatif</span><span class="sxs-lookup"><span data-stu-id="c3298-120">optional</span></span>|<span data-ttu-id="c3298-121">Spécifie la version minimale de l’ensemble d’API requis par votre complément.</span><span class="sxs-lookup"><span data-stu-id="c3298-121">Specifies the minimum version of the API set required by your add-in.</span></span> <span data-ttu-id="c3298-122">Remplace la valeur de **DefaultMinVersion**, si elle est spécifiée dans l’élément parent [sets](sets.md) .</span><span class="sxs-lookup"><span data-stu-id="c3298-122">Overrides the value of **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="c3298-123">Remarques</span><span class="sxs-lookup"><span data-stu-id="c3298-123">Remarks</span></span>

<span data-ttu-id="c3298-124">Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="c3298-124">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="c3298-125">Pour plus d’informations sur l’attribut **MinVersion** de l’élément **Set** et sur l’attribut **DefaultMinVersion** de l’élément **sets** , voir [Set the requirements ELEMENT dans le manifeste](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="c3298-125">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c3298-126">Pour les compléments de messagerie, il n'existe qu’un seul `"Mailbox"`ensemble de conditions requises disponible.</span><span class="sxs-lookup"><span data-stu-id="c3298-126">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="c3298-127">Cet ensemble de conditions requises contient le sous-ensemble complet de l’API pris en charge dans les compléments de messagerie pour Outlook, et vous devez spécifier `"Mailbox"`l’ensemble de conditions requises dans le manifeste de votre complément de messagerie (ce n’est pas facultatif, comme c’est le cas pour les compléments de contenu et de volet Office). </span><span class="sxs-lookup"><span data-stu-id="c3298-127">This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins).</span></span> <span data-ttu-id="c3298-128">De même, vous ne pouvez pas déclarer une prise en charge pour des méthodes spécifiques dans les compléments de messagerie.</span><span class="sxs-lookup"><span data-stu-id="c3298-128">Also, you can't declare support for specific methods in mail add-ins.</span></span>
