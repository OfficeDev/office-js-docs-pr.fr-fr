---
title: Élément SupportUrl dans le fichier manifest
description: L’élément SupportUrl spécifie l’URL d’une page qui fournit des informations de prise en charge pour votre complément.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: e38030062c48936f925126e896cd74e660164a5d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720343"
---
# <a name="supporturl-element"></a><span data-ttu-id="b415b-103">SupportUrl, élément</span><span class="sxs-lookup"><span data-stu-id="b415b-103">SupportUrl element</span></span>

<span data-ttu-id="b415b-104">Spécifie l’URL d’une page qui fournit des informations de prise en charge pour votre complément.</span><span class="sxs-lookup"><span data-stu-id="b415b-104">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="b415b-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="b415b-105">Syntax</span></span>

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="b415b-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="b415b-106">Contained in</span></span>

[<span data-ttu-id="b415b-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="b415b-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="b415b-108">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="b415b-108">Can contain</span></span>

|  <span data-ttu-id="b415b-109">Élément</span><span class="sxs-lookup"><span data-stu-id="b415b-109">Element</span></span> | <span data-ttu-id="b415b-110">Requis</span><span class="sxs-lookup"><span data-stu-id="b415b-110">Required</span></span> | <span data-ttu-id="b415b-111">Description</span><span class="sxs-lookup"><span data-stu-id="b415b-111">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="b415b-112">Override</span><span class="sxs-lookup"><span data-stu-id="b415b-112">Override</span></span>](override.md)   | <span data-ttu-id="b415b-113">Non</span><span class="sxs-lookup"><span data-stu-id="b415b-113">No</span></span> | <span data-ttu-id="b415b-114">Spécifie le paramètre pour les URL de paramètres régionaux supplémentaires</span><span class="sxs-lookup"><span data-stu-id="b415b-114">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="b415b-115">Attributs</span><span class="sxs-lookup"><span data-stu-id="b415b-115">Attributes</span></span>

|<span data-ttu-id="b415b-116">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="b415b-116">**Attribute**</span></span>|<span data-ttu-id="b415b-117">**Type**</span><span class="sxs-lookup"><span data-stu-id="b415b-117">**Type**</span></span>|<span data-ttu-id="b415b-118">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="b415b-118">**Required**</span></span>|<span data-ttu-id="b415b-119">**Description**</span><span class="sxs-lookup"><span data-stu-id="b415b-119">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="b415b-120">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="b415b-120">DefaultValue</span></span>|<span data-ttu-id="b415b-121">URL</span><span class="sxs-lookup"><span data-stu-id="b415b-121">URL</span></span>|<span data-ttu-id="b415b-122">obligatoire</span><span class="sxs-lookup"><span data-stu-id="b415b-122">required</span></span>|<span data-ttu-id="b415b-123">Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="b415b-123">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
