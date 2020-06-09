---
title: Élément SupportUrl dans le fichier manifest
description: L’élément SupportUrl spécifie l’URL d’une page qui fournit des informations de prise en charge pour votre complément.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f75ee811699823a501ac594e66daaaf3f93c2782
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608704"
---
# <a name="supporturl-element"></a><span data-ttu-id="4707e-103">SupportUrl, élément</span><span class="sxs-lookup"><span data-stu-id="4707e-103">SupportUrl element</span></span>

<span data-ttu-id="4707e-104">Spécifie l’URL d’une page qui fournit des informations de prise en charge pour votre complément.</span><span class="sxs-lookup"><span data-stu-id="4707e-104">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="4707e-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="4707e-105">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="4707e-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="4707e-106">Contained in</span></span>

[<span data-ttu-id="4707e-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="4707e-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="4707e-108">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="4707e-108">Can contain</span></span>

|  <span data-ttu-id="4707e-109">Élément</span><span class="sxs-lookup"><span data-stu-id="4707e-109">Element</span></span> | <span data-ttu-id="4707e-110">Requis</span><span class="sxs-lookup"><span data-stu-id="4707e-110">Required</span></span> | <span data-ttu-id="4707e-111">Description</span><span class="sxs-lookup"><span data-stu-id="4707e-111">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="4707e-112">Override</span><span class="sxs-lookup"><span data-stu-id="4707e-112">Override</span></span>](override.md)   | <span data-ttu-id="4707e-113">Non</span><span class="sxs-lookup"><span data-stu-id="4707e-113">No</span></span> | <span data-ttu-id="4707e-114">Spécifie le paramètre pour les URL de paramètres régionaux supplémentaires</span><span class="sxs-lookup"><span data-stu-id="4707e-114">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="4707e-115">Attributs</span><span class="sxs-lookup"><span data-stu-id="4707e-115">Attributes</span></span>

|<span data-ttu-id="4707e-116">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="4707e-116">**Attribute**</span></span>|<span data-ttu-id="4707e-117">**Type**</span><span class="sxs-lookup"><span data-stu-id="4707e-117">**Type**</span></span>|<span data-ttu-id="4707e-118">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="4707e-118">**Required**</span></span>|<span data-ttu-id="4707e-119">**Description**</span><span class="sxs-lookup"><span data-stu-id="4707e-119">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="4707e-120">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="4707e-120">DefaultValue</span></span>|<span data-ttu-id="4707e-121">URL</span><span class="sxs-lookup"><span data-stu-id="4707e-121">URL</span></span>|<span data-ttu-id="4707e-122">obligatoire</span><span class="sxs-lookup"><span data-stu-id="4707e-122">required</span></span>|<span data-ttu-id="4707e-123">Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="4707e-123">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
