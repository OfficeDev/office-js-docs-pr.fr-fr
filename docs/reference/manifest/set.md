---
title: Élément Set dans le fichier manifeste
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d86b3123ff856e8618f92629308787b543e8228b
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324805"
---
# <a name="set-element"></a><span data-ttu-id="b81e0-102">Élément Set</span><span class="sxs-lookup"><span data-stu-id="b81e0-102">Set element</span></span>

<span data-ttu-id="b81e0-103">Spécifie un ensemble de conditions requises de l’API JavaScript Office que votre complément Office doit activer.</span><span class="sxs-lookup"><span data-stu-id="b81e0-103">Specifies a requirement set from the Office JavaScript API that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="b81e0-104">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="b81e0-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b81e0-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="b81e0-105">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="b81e0-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="b81e0-106">Contained in</span></span>

[<span data-ttu-id="b81e0-107">Ensembles</span><span class="sxs-lookup"><span data-stu-id="b81e0-107">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="b81e0-108">Attributs</span><span class="sxs-lookup"><span data-stu-id="b81e0-108">Attributes</span></span>

|<span data-ttu-id="b81e0-109">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="b81e0-109">**Attribute**</span></span>|<span data-ttu-id="b81e0-110">**Type**</span><span class="sxs-lookup"><span data-stu-id="b81e0-110">**Type**</span></span>|<span data-ttu-id="b81e0-111">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="b81e0-111">**Required**</span></span>|<span data-ttu-id="b81e0-112">**Description**</span><span class="sxs-lookup"><span data-stu-id="b81e0-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="b81e0-113">Nom</span><span class="sxs-lookup"><span data-stu-id="b81e0-113">Name</span></span>|<span data-ttu-id="b81e0-114">string</span><span class="sxs-lookup"><span data-stu-id="b81e0-114">string</span></span>|<span data-ttu-id="b81e0-115">obligatoire</span><span class="sxs-lookup"><span data-stu-id="b81e0-115">required</span></span>|<span data-ttu-id="b81e0-116">Nom d’un [ensemble de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="b81e0-116">The name of a [requirement set](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>|
|<span data-ttu-id="b81e0-117">MinVersion</span><span class="sxs-lookup"><span data-stu-id="b81e0-117">MinVersion</span></span>|<span data-ttu-id="b81e0-118">chaîne</span><span class="sxs-lookup"><span data-stu-id="b81e0-118">string</span></span>|<span data-ttu-id="b81e0-119">facultatif</span><span class="sxs-lookup"><span data-stu-id="b81e0-119">optional</span></span>|<span data-ttu-id="b81e0-120">Spécifie la version minimale de l’ensemble d’API requis par votre complément.</span><span class="sxs-lookup"><span data-stu-id="b81e0-120">Specifies the minimum version of the API set required by your add-in.</span></span> <span data-ttu-id="b81e0-121">Remplace la valeur de **DefaultMinVersion**, si elle est spécifiée dans l’élément parent [sets](sets.md) .</span><span class="sxs-lookup"><span data-stu-id="b81e0-121">Overrides the value of **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="b81e0-122">Remarques</span><span class="sxs-lookup"><span data-stu-id="b81e0-122">Remarks</span></span>

<span data-ttu-id="b81e0-123">Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="b81e0-123">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="b81e0-124">Pour plus d’informations sur l’attribut **MinVersion** de l’élément **Set** et sur l’attribut **DefaultMinVersion** de l’élément **sets** , voir [Set the requirements ELEMENT dans le manifeste](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="b81e0-124">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="b81e0-125">Pour les compléments de messagerie, il n'existe qu’un seul `"Mailbox"`ensemble de conditions requises disponible.</span><span class="sxs-lookup"><span data-stu-id="b81e0-125">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="b81e0-126">Cet ensemble de conditions requises contient le sous-ensemble complet de l’API pris en charge dans les compléments de messagerie pour Outlook, et vous devez spécifier `"Mailbox"`l’ensemble de conditions requises dans le manifeste de votre complément de messagerie (ce n’est pas facultatif, comme c’est le cas pour les compléments de contenu et de volet Office). </span><span class="sxs-lookup"><span data-stu-id="b81e0-126">This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins).</span></span> <span data-ttu-id="b81e0-127">De même, vous ne pouvez pas déclarer une prise en charge pour des méthodes spécifiques dans les compléments de messagerie.</span><span class="sxs-lookup"><span data-stu-id="b81e0-127">Also, you can't declare support for specific methods in mail add-ins.</span></span>
