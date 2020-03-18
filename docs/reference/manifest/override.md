---
title: Élément Override dans le fichier manifest
description: L’élément override vous permet de spécifier la valeur d’un paramètre pour des paramètres régionaux supplémentaires.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 39e706dc981d405fcfcc508626578f34931efbcb
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718026"
---
# <a name="override-element"></a><span data-ttu-id="84c03-103">Élément Override</span><span class="sxs-lookup"><span data-stu-id="84c03-103">Override element</span></span>

<span data-ttu-id="84c03-104">Fournit une manière de spécifier la valeur d’un paramètre pour d’autres paramètres régionaux.</span><span class="sxs-lookup"><span data-stu-id="84c03-104">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="84c03-105">**Type de complément:** application de contenu, de volet Office, de messagerie (Mail)</span><span class="sxs-lookup"><span data-stu-id="84c03-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="84c03-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="84c03-106">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="84c03-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="84c03-107">Contained in</span></span>

|<span data-ttu-id="84c03-108">**Élément**</span><span class="sxs-lookup"><span data-stu-id="84c03-108">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="84c03-109">CitationText</span><span class="sxs-lookup"><span data-stu-id="84c03-109">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="84c03-110">Description</span><span class="sxs-lookup"><span data-stu-id="84c03-110">Description</span></span>](description.md)|
|[<span data-ttu-id="84c03-111">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="84c03-111">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="84c03-112">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="84c03-112">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="84c03-113">DisplayName</span><span class="sxs-lookup"><span data-stu-id="84c03-113">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="84c03-114">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="84c03-114">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="84c03-115">IconUrl</span><span class="sxs-lookup"><span data-stu-id="84c03-115">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="84c03-116">QueryUri</span><span class="sxs-lookup"><span data-stu-id="84c03-116">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="84c03-117">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="84c03-117">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="84c03-118">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="84c03-118">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="84c03-119">Attributs</span><span class="sxs-lookup"><span data-stu-id="84c03-119">Attributes</span></span>

|<span data-ttu-id="84c03-120">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="84c03-120">**Attribute**</span></span>|<span data-ttu-id="84c03-121">**Type**</span><span class="sxs-lookup"><span data-stu-id="84c03-121">**Type**</span></span>|<span data-ttu-id="84c03-122">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="84c03-122">**Required**</span></span>|<span data-ttu-id="84c03-123">**Description**</span><span class="sxs-lookup"><span data-stu-id="84c03-123">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="84c03-124">Paramètres régionaux</span><span class="sxs-lookup"><span data-stu-id="84c03-124">Locale</span></span>|<span data-ttu-id="84c03-125">string</span><span class="sxs-lookup"><span data-stu-id="84c03-125">string</span></span>|<span data-ttu-id="84c03-126">obligatoire</span><span class="sxs-lookup"><span data-stu-id="84c03-126">required</span></span>|<span data-ttu-id="84c03-127">Spécifie le nom de culture des paramètres régionaux pour ce remplacement au format de balise de langue BCP 47, comme `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="84c03-127">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="84c03-128">Valeur</span><span class="sxs-lookup"><span data-stu-id="84c03-128">Value</span></span>|<span data-ttu-id="84c03-129">string</span><span class="sxs-lookup"><span data-stu-id="84c03-129">string</span></span>|<span data-ttu-id="84c03-130">obligatoire</span><span class="sxs-lookup"><span data-stu-id="84c03-130">required</span></span>|<span data-ttu-id="84c03-131">Spécifie la valeur du paramètre exprimée pour les paramètres régionaux spécifiés.</span><span class="sxs-lookup"><span data-stu-id="84c03-131">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="84c03-132">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="84c03-132">See also</span></span>

- [<span data-ttu-id="84c03-133">Localisation des compléments Office</span><span class="sxs-lookup"><span data-stu-id="84c03-133">Localization for Office Add-ins</span></span>](../../develop/localization.md)
