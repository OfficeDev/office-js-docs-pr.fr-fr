---
title: Élément Set dans le fichier manifeste
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 0f137f7b08d6f1d0b0d972173c8085713b0f979d
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432766"
---
# <a name="set-element"></a><span data-ttu-id="c47ac-102">Élément Set</span><span class="sxs-lookup"><span data-stu-id="c47ac-102">Set element</span></span>

<span data-ttu-id="c47ac-103">Spécifie un ensemble de conditions requises de l’API JavaScript pour Office nécessaires à l’activation de votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="c47ac-103">Specifies a requirement set from the JavaScript API for Office that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="c47ac-104">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="c47ac-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="c47ac-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="c47ac-105">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="c47ac-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="c47ac-106">Contained in</span></span>

[<span data-ttu-id="c47ac-107">Ensembles</span><span class="sxs-lookup"><span data-stu-id="c47ac-107">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="c47ac-108">Attributs</span><span class="sxs-lookup"><span data-stu-id="c47ac-108">Attributes</span></span>

|<span data-ttu-id="c47ac-109">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="c47ac-109">**Attribute**</span></span>|<span data-ttu-id="c47ac-110">**Type**</span><span class="sxs-lookup"><span data-stu-id="c47ac-110">**Type**</span></span>|<span data-ttu-id="c47ac-111">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="c47ac-111">**Required**</span></span>|<span data-ttu-id="c47ac-112">**Description**</span><span class="sxs-lookup"><span data-stu-id="c47ac-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="c47ac-113">Nom</span><span class="sxs-lookup"><span data-stu-id="c47ac-113">Name</span></span>|<span data-ttu-id="c47ac-114">string</span><span class="sxs-lookup"><span data-stu-id="c47ac-114">string</span></span>|<span data-ttu-id="c47ac-115">obligatoire</span><span class="sxs-lookup"><span data-stu-id="c47ac-115">required</span></span>|<span data-ttu-id="c47ac-116">Nom d’un [ensemble de conditions requises](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="c47ac-116">The name of a [requirement set](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>|
|<span data-ttu-id="c47ac-117">MinVersion</span><span class="sxs-lookup"><span data-stu-id="c47ac-117">MinVersion</span></span>|<span data-ttu-id="c47ac-118">chaîne</span><span class="sxs-lookup"><span data-stu-id="c47ac-118">string</span></span>|<span data-ttu-id="c47ac-119">facultatif</span><span class="sxs-lookup"><span data-stu-id="c47ac-119">optional</span></span>|<span data-ttu-id="c47ac-p101">Spécifie la version minimale de l’ensemble d’API requis par votre complément. Remplace la valeur de **DefaultMinVersion**, si elle est spécifiée dans l’élément parent [Sets](sets.md).</span><span class="sxs-lookup"><span data-stu-id="c47ac-p101">Specifies the minimum version of the API set required by your add-in. Overrides the value of  **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="c47ac-122">Remarques</span><span class="sxs-lookup"><span data-stu-id="c47ac-122">Remarks</span></span>

<span data-ttu-id="c47ac-123">Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="c47ac-123">For more information, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="c47ac-124">Pour plus d'informations sur l’attribut **MinVersion** de l’élément **Set** et sur l’attribut **DefaultMinVersion** de l’élément **Sets**, voir l’article relatif à la [définition de l’élément Requirements dans le manifeste](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="c47ac-124">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="c47ac-125">Pour les compléments de messagerie, il n'existe qu’un seul `"Mailbox"`ensemble de conditions requises disponible.</span><span class="sxs-lookup"><span data-stu-id="c47ac-125">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="c47ac-126">Cet ensemble de conditions requises contient le sous-ensemble complet de l’API pris en charge dans les compléments de messagerie pour Outlook, et vous devez spécifier `"Mailbox"`l’ensemble de conditions requises dans le manifeste de votre complément de messagerie (ce n’est pas facultatif, comme c’est le cas pour les compléments de contenu et de volet Office). </span><span class="sxs-lookup"><span data-stu-id="c47ac-126">Important  For mail add-ins, there is only one   requirement set available. This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins). Also, you can't declare support for specific methods in mail add-ins.</span></span> <span data-ttu-id="c47ac-127">De même, vous ne pouvez pas déclarer une prise en charge pour des méthodes spécifiques dans les compléments de messagerie.</span><span class="sxs-lookup"><span data-stu-id="c47ac-127">Also, you can't declare support for specific methods in mail add-ins.</span></span>
