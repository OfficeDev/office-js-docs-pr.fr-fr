---
title: Élément Override dans le fichier manifest
description: L’élément override vous permet de spécifier la valeur d’un paramètre pour des paramètres régionaux supplémentaires.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 139a4089a36d8a8adfa71d4a0947b02f5b163b52
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641451"
---
# <a name="override-element"></a><span data-ttu-id="652ec-103">Élément Override</span><span class="sxs-lookup"><span data-stu-id="652ec-103">Override element</span></span>

<span data-ttu-id="652ec-104">Fournit une manière de spécifier la valeur d’un paramètre pour d’autres paramètres régionaux.</span><span class="sxs-lookup"><span data-stu-id="652ec-104">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="652ec-105">**Type de complément:** application de contenu, de volet Office, de messagerie (Mail)</span><span class="sxs-lookup"><span data-stu-id="652ec-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="652ec-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="652ec-106">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="652ec-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="652ec-107">Contained in</span></span>

|<span data-ttu-id="652ec-108">Élément</span><span class="sxs-lookup"><span data-stu-id="652ec-108">Element</span></span>|
|:-----|
|[<span data-ttu-id="652ec-109">CitationText</span><span class="sxs-lookup"><span data-stu-id="652ec-109">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="652ec-110">Description</span><span class="sxs-lookup"><span data-stu-id="652ec-110">Description</span></span>](description.md)|
|[<span data-ttu-id="652ec-111">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="652ec-111">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="652ec-112">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="652ec-112">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="652ec-113">DisplayName</span><span class="sxs-lookup"><span data-stu-id="652ec-113">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="652ec-114">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="652ec-114">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="652ec-115">IconUrl</span><span class="sxs-lookup"><span data-stu-id="652ec-115">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="652ec-116">QueryUri</span><span class="sxs-lookup"><span data-stu-id="652ec-116">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="652ec-117">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="652ec-117">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="652ec-118">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="652ec-118">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="652ec-119">Attributs</span><span class="sxs-lookup"><span data-stu-id="652ec-119">Attributes</span></span>

|<span data-ttu-id="652ec-120">Attribut</span><span class="sxs-lookup"><span data-stu-id="652ec-120">Attribute</span></span>|<span data-ttu-id="652ec-121">Type</span><span class="sxs-lookup"><span data-stu-id="652ec-121">Type</span></span>|<span data-ttu-id="652ec-122">Requis</span><span class="sxs-lookup"><span data-stu-id="652ec-122">Required</span></span>|<span data-ttu-id="652ec-123">Description</span><span class="sxs-lookup"><span data-stu-id="652ec-123">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="652ec-124">Paramètres régionaux</span><span class="sxs-lookup"><span data-stu-id="652ec-124">Locale</span></span>|<span data-ttu-id="652ec-125">string</span><span class="sxs-lookup"><span data-stu-id="652ec-125">string</span></span>|<span data-ttu-id="652ec-126">obligatoire</span><span class="sxs-lookup"><span data-stu-id="652ec-126">required</span></span>|<span data-ttu-id="652ec-127">Spécifie le nom de culture des paramètres régionaux pour ce remplacement au format de balise de langue BCP 47, comme `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="652ec-127">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="652ec-128">Valeur</span><span class="sxs-lookup"><span data-stu-id="652ec-128">Value</span></span>|<span data-ttu-id="652ec-129">string</span><span class="sxs-lookup"><span data-stu-id="652ec-129">string</span></span>|<span data-ttu-id="652ec-130">obligatoire</span><span class="sxs-lookup"><span data-stu-id="652ec-130">required</span></span>|<span data-ttu-id="652ec-131">Spécifie la valeur du paramètre exprimée pour les paramètres régionaux spécifiés.</span><span class="sxs-lookup"><span data-stu-id="652ec-131">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="652ec-132">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="652ec-132">See also</span></span>

- [<span data-ttu-id="652ec-133">Localisation des compléments Office</span><span class="sxs-lookup"><span data-stu-id="652ec-133">Localization for Office Add-ins</span></span>](../../develop/localization.md)
