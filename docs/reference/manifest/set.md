---
title: Élément Set dans le fichier manifeste
description: L’élément Set spécifie un ensemble de conditions requises de l’API JavaScript pour Office requis pour l’activation de votre complément Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: e9a70da0dc38c3aee077eb5e7f47cdf8e6dc2d32
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717914"
---
# <a name="set-element"></a><span data-ttu-id="c91b4-103">Élément Set</span><span class="sxs-lookup"><span data-stu-id="c91b4-103">Set element</span></span>

<span data-ttu-id="c91b4-104">Spécifie un ensemble de conditions requises de l’API JavaScript Office que votre complément Office doit activer.</span><span class="sxs-lookup"><span data-stu-id="c91b4-104">Specifies a requirement set from the Office JavaScript API that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="c91b4-105">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="c91b4-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="c91b4-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="c91b4-106">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="c91b4-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="c91b4-107">Contained in</span></span>

[<span data-ttu-id="c91b4-108">Ensembles</span><span class="sxs-lookup"><span data-stu-id="c91b4-108">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="c91b4-109">Attributs</span><span class="sxs-lookup"><span data-stu-id="c91b4-109">Attributes</span></span>

|<span data-ttu-id="c91b4-110">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="c91b4-110">**Attribute**</span></span>|<span data-ttu-id="c91b4-111">**Type**</span><span class="sxs-lookup"><span data-stu-id="c91b4-111">**Type**</span></span>|<span data-ttu-id="c91b4-112">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="c91b4-112">**Required**</span></span>|<span data-ttu-id="c91b4-113">**Description**</span><span class="sxs-lookup"><span data-stu-id="c91b4-113">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="c91b4-114">Nom</span><span class="sxs-lookup"><span data-stu-id="c91b4-114">Name</span></span>|<span data-ttu-id="c91b4-115">string</span><span class="sxs-lookup"><span data-stu-id="c91b4-115">string</span></span>|<span data-ttu-id="c91b4-116">obligatoire</span><span class="sxs-lookup"><span data-stu-id="c91b4-116">required</span></span>|<span data-ttu-id="c91b4-117">Nom d’un [ensemble de conditions requises](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="c91b4-117">The name of a [requirement set](../../develop/office-versions-and-requirement-sets.md).</span></span>|
|<span data-ttu-id="c91b4-118">MinVersion</span><span class="sxs-lookup"><span data-stu-id="c91b4-118">MinVersion</span></span>|<span data-ttu-id="c91b4-119">chaîne</span><span class="sxs-lookup"><span data-stu-id="c91b4-119">string</span></span>|<span data-ttu-id="c91b4-120">facultatif</span><span class="sxs-lookup"><span data-stu-id="c91b4-120">optional</span></span>|<span data-ttu-id="c91b4-121">Spécifie la version minimale de l’ensemble d’API requis par votre complément.</span><span class="sxs-lookup"><span data-stu-id="c91b4-121">Specifies the minimum version of the API set required by your add-in.</span></span> <span data-ttu-id="c91b4-122">Remplace la valeur de **DefaultMinVersion**, si elle est spécifiée dans l’élément parent [sets](sets.md) .</span><span class="sxs-lookup"><span data-stu-id="c91b4-122">Overrides the value of **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="c91b4-123">Remarques</span><span class="sxs-lookup"><span data-stu-id="c91b4-123">Remarks</span></span>

<span data-ttu-id="c91b4-124">Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="c91b4-124">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="c91b4-125">Pour plus d’informations sur l’attribut **MinVersion** de l’élément **Set** et sur l’attribut **DefaultMinVersion** de l’élément **sets** , voir [Set the requirements ELEMENT dans le manifeste](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="c91b4-125">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="c91b4-126">Pour les compléments de messagerie, il n'existe qu’un seul `"Mailbox"`ensemble de conditions requises disponible.</span><span class="sxs-lookup"><span data-stu-id="c91b4-126">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="c91b4-127">Cet ensemble de conditions requises contient le sous-ensemble complet de l’API pris en charge dans les compléments de messagerie pour Outlook, et vous devez spécifier `"Mailbox"`l’ensemble de conditions requises dans le manifeste de votre complément de messagerie (ce n’est pas facultatif, comme c’est le cas pour les compléments de contenu et de volet Office). </span><span class="sxs-lookup"><span data-stu-id="c91b4-127">This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins).</span></span> <span data-ttu-id="c91b4-128">De même, vous ne pouvez pas déclarer une prise en charge pour des méthodes spécifiques dans les compléments de messagerie.</span><span class="sxs-lookup"><span data-stu-id="c91b4-128">Also, you can't declare support for specific methods in mail add-ins.</span></span>
