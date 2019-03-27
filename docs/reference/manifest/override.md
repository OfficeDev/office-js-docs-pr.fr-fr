---
title: Élément Override dans le fichier manifest
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 020ae490dacbb9b8c493dc022c23d0ebf311a1b9
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870057"
---
# <a name="override-element"></a><span data-ttu-id="3d014-102">Élément Override</span><span class="sxs-lookup"><span data-stu-id="3d014-102">Override element</span></span>

<span data-ttu-id="3d014-103">Fournit une manière de spécifier la valeur d’un paramètre pour d’autres paramètres régionaux.</span><span class="sxs-lookup"><span data-stu-id="3d014-103">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="3d014-104">**Type de complément:** application de contenu, de volet Office, de messagerie (Mail)</span><span class="sxs-lookup"><span data-stu-id="3d014-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="3d014-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="3d014-105">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="3d014-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="3d014-106">Contained in</span></span>

|<span data-ttu-id="3d014-107">**Élément**</span><span class="sxs-lookup"><span data-stu-id="3d014-107">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="3d014-108">CitationText</span><span class="sxs-lookup"><span data-stu-id="3d014-108">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="3d014-109">Description</span><span class="sxs-lookup"><span data-stu-id="3d014-109">Description</span></span>](description.md)|
|[<span data-ttu-id="3d014-110">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="3d014-110">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="3d014-111">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="3d014-111">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="3d014-112">DisplayName</span><span class="sxs-lookup"><span data-stu-id="3d014-112">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="3d014-113">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="3d014-113">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="3d014-114">IconUrl</span><span class="sxs-lookup"><span data-stu-id="3d014-114">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="3d014-115">QueryUri</span><span class="sxs-lookup"><span data-stu-id="3d014-115">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="3d014-116">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="3d014-116">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="3d014-117">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="3d014-117">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="3d014-118">Attributs</span><span class="sxs-lookup"><span data-stu-id="3d014-118">Attributes</span></span>

|<span data-ttu-id="3d014-119">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="3d014-119">**Attribute**</span></span>|<span data-ttu-id="3d014-120">**Type**</span><span class="sxs-lookup"><span data-stu-id="3d014-120">**Type**</span></span>|<span data-ttu-id="3d014-121">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="3d014-121">**Required**</span></span>|<span data-ttu-id="3d014-122">**Description**</span><span class="sxs-lookup"><span data-stu-id="3d014-122">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="3d014-123">Paramètres régionaux</span><span class="sxs-lookup"><span data-stu-id="3d014-123">Locale</span></span>|<span data-ttu-id="3d014-124">string</span><span class="sxs-lookup"><span data-stu-id="3d014-124">string</span></span>|<span data-ttu-id="3d014-125">obligatoire</span><span class="sxs-lookup"><span data-stu-id="3d014-125">required</span></span>|<span data-ttu-id="3d014-126">Spécifie le nom de culture des paramètres régionaux pour ce remplacement au format de balise de langue BCP 47, comme `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="3d014-126">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="3d014-127">Valeur</span><span class="sxs-lookup"><span data-stu-id="3d014-127">Value</span></span>|<span data-ttu-id="3d014-128">string</span><span class="sxs-lookup"><span data-stu-id="3d014-128">string</span></span>|<span data-ttu-id="3d014-129">obligatoire</span><span class="sxs-lookup"><span data-stu-id="3d014-129">required</span></span>|<span data-ttu-id="3d014-130">Spécifie la valeur du paramètre exprimée pour les paramètres régionaux spécifiés.</span><span class="sxs-lookup"><span data-stu-id="3d014-130">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="3d014-131">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="3d014-131">See also</span></span>

- [<span data-ttu-id="3d014-132">Localisation des compléments Office</span><span class="sxs-lookup"><span data-stu-id="3d014-132">Localization for Office Add-ins</span></span>](/office/dev/add-ins/develop/localization)
    
