---
title: Élément Sets dans le fichier manifeste
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: b7e78ae05f8409f38c885a1d6a328347d00d0df1
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433655"
---
# <a name="sets-element"></a><span data-ttu-id="ce857-102">Sets, élément</span><span class="sxs-lookup"><span data-stu-id="ce857-102">Sets element</span></span>

<span data-ttu-id="ce857-103">Spécifie le sous-ensemble minimal de l’API JavaScript pour Office nécessaire à l’activation de votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="ce857-103">Specifies the minimum subset of the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="ce857-104">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="ce857-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="ce857-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="ce857-105">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="ce857-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="ce857-106">Contained in</span></span>

[<span data-ttu-id="ce857-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ce857-107">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="ce857-108">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="ce857-108">Can contain</span></span>

[<span data-ttu-id="ce857-109">Ensemble</span><span class="sxs-lookup"><span data-stu-id="ce857-109">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="ce857-110">Attributs</span><span class="sxs-lookup"><span data-stu-id="ce857-110">Attributes</span></span>

|<span data-ttu-id="ce857-111">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="ce857-111">**Attribute**</span></span>|<span data-ttu-id="ce857-112">**Type**</span><span class="sxs-lookup"><span data-stu-id="ce857-112">**Type**</span></span>|<span data-ttu-id="ce857-113">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="ce857-113">**Required**</span></span>|<span data-ttu-id="ce857-114">**Description**</span><span class="sxs-lookup"><span data-stu-id="ce857-114">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="ce857-115">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="ce857-115">DefaultMinVersion</span></span>|<span data-ttu-id="ce857-116">chaîne</span><span class="sxs-lookup"><span data-stu-id="ce857-116">string</span></span>|<span data-ttu-id="ce857-117">facultatif</span><span class="sxs-lookup"><span data-stu-id="ce857-117">optional</span></span>|<span data-ttu-id="ce857-p101">Spécifie la valeur de l’attribut **MinVersion** par défaut pour tous les éléments [Set](set.md) enfants. La valeur par défaut est « 1.1 ».</span><span class="sxs-lookup"><span data-stu-id="ce857-p101">Specifies the default  **MinVersion** attribute value for all child [Set](set.md) elements. The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="ce857-120">Remarques</span><span class="sxs-lookup"><span data-stu-id="ce857-120">Remarks</span></span>

<span data-ttu-id="ce857-121">Pour plus d’informations concernant les ensembles de conditions requises, voir [Versions d’Office et ensembles de conditions requises](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="ce857-121">For more information, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="ce857-122">Pour plus d'informations sur l’attribut **MinVersion** de l’élément **Set** et sur l’attribut **DefaultMinVersion** de l’élément **Sets**, voir l’article relatif à la [définition de l’élément Requirements dans le manifeste](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="ce857-122">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

