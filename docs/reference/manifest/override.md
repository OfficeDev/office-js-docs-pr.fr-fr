---
title: Élément Override dans le fichier manifest
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: d1d2400312f12116b1ac5f4010135541e783dcc7
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432864"
---
# <a name="override-element"></a><span data-ttu-id="228a5-102">Élément Override</span><span class="sxs-lookup"><span data-stu-id="228a5-102">Override element</span></span>

<span data-ttu-id="228a5-103">Fournit une manière de spécifier la valeur d’un paramètre pour d’autres paramètres régionaux.</span><span class="sxs-lookup"><span data-stu-id="228a5-103">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="228a5-104">**Type de complément:** application de contenu, de volet Office, de messagerie (Mail)</span><span class="sxs-lookup"><span data-stu-id="228a5-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="228a5-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="228a5-105">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="228a5-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="228a5-106">Contained in</span></span>

|<span data-ttu-id="228a5-107">**Élément**</span><span class="sxs-lookup"><span data-stu-id="228a5-107">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="228a5-108">CitationText</span><span class="sxs-lookup"><span data-stu-id="228a5-108">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="228a5-109">Description</span><span class="sxs-lookup"><span data-stu-id="228a5-109">Description</span></span>](description.md)|
|[<span data-ttu-id="228a5-110">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="228a5-110">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="228a5-111">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="228a5-111">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="228a5-112">DisplayName</span><span class="sxs-lookup"><span data-stu-id="228a5-112">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="228a5-113">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="228a5-113">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="228a5-114">IconUrl</span><span class="sxs-lookup"><span data-stu-id="228a5-114">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="228a5-115">QueryUri</span><span class="sxs-lookup"><span data-stu-id="228a5-115">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="228a5-116">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="228a5-116">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="228a5-117">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="228a5-117">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="228a5-118">Attributs</span><span class="sxs-lookup"><span data-stu-id="228a5-118">Attributes</span></span>

|<span data-ttu-id="228a5-119">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="228a5-119">**Attribute**</span></span>|<span data-ttu-id="228a5-120">**Type**</span><span class="sxs-lookup"><span data-stu-id="228a5-120">**Type**</span></span>|<span data-ttu-id="228a5-121">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="228a5-121">**Required**</span></span>|<span data-ttu-id="228a5-122">**Description**</span><span class="sxs-lookup"><span data-stu-id="228a5-122">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="228a5-123">Locale</span><span class="sxs-lookup"><span data-stu-id="228a5-123">Locale</span></span>|<span data-ttu-id="228a5-124">string</span><span class="sxs-lookup"><span data-stu-id="228a5-124">string</span></span>|<span data-ttu-id="228a5-125">obligatoire</span><span class="sxs-lookup"><span data-stu-id="228a5-125">required</span></span>|<span data-ttu-id="228a5-126">Spécifie le nom de culture des paramètres régionaux pour ce remplacement au format de balise de langue BCP 47, comme `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="228a5-126">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="228a5-127">Valeur</span><span class="sxs-lookup"><span data-stu-id="228a5-127">Value</span></span>|<span data-ttu-id="228a5-128">string</span><span class="sxs-lookup"><span data-stu-id="228a5-128">string</span></span>|<span data-ttu-id="228a5-129">obligatoire</span><span class="sxs-lookup"><span data-stu-id="228a5-129">required</span></span>|<span data-ttu-id="228a5-130">Spécifie la valeur du paramètre exprimée pour les paramètres régionaux spécifiés.</span><span class="sxs-lookup"><span data-stu-id="228a5-130">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="228a5-131">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="228a5-131">See also</span></span>

- [<span data-ttu-id="228a5-132">Localisation des compléments Office</span><span class="sxs-lookup"><span data-stu-id="228a5-132">Localization for Office Add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/localization)
    
