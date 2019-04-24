---
title: Élément SupportUrl dans le fichier manifest
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 18b9b7c4df9def70ab42ae213066188ac04c07a7
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450414"
---
# <a name="supporturl-element"></a><span data-ttu-id="f21c2-102">Élément SupportUrl</span><span class="sxs-lookup"><span data-stu-id="f21c2-102">SupportUrl element</span></span>

<span data-ttu-id="f21c2-103">Spécifie l’URL d’une page qui fournit des informations de prise en charge pour votre complément.</span><span class="sxs-lookup"><span data-stu-id="f21c2-103">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="f21c2-104">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="f21c2-104">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="f21c2-105">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="f21c2-105">Contained in</span></span>

[<span data-ttu-id="f21c2-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="f21c2-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="f21c2-107">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="f21c2-107">Can contain</span></span>

|  <span data-ttu-id="f21c2-108">Élément</span><span class="sxs-lookup"><span data-stu-id="f21c2-108">Element</span></span> | <span data-ttu-id="f21c2-109">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="f21c2-109">Required</span></span> | <span data-ttu-id="f21c2-110">Description</span><span class="sxs-lookup"><span data-stu-id="f21c2-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="f21c2-111">Override</span><span class="sxs-lookup"><span data-stu-id="f21c2-111">Override</span></span>](override.md)   | <span data-ttu-id="f21c2-112">Non</span><span class="sxs-lookup"><span data-stu-id="f21c2-112">No</span></span> | <span data-ttu-id="f21c2-113">Spécifie le paramètre pour les URL de paramètres régionaux supplémentaires</span><span class="sxs-lookup"><span data-stu-id="f21c2-113">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="f21c2-114">Attributs</span><span class="sxs-lookup"><span data-stu-id="f21c2-114">Attributes</span></span>

|<span data-ttu-id="f21c2-115">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="f21c2-115">**Attribute**</span></span>|<span data-ttu-id="f21c2-116">**Type**</span><span class="sxs-lookup"><span data-stu-id="f21c2-116">**Type**</span></span>|<span data-ttu-id="f21c2-117">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="f21c2-117">**Required**</span></span>|<span data-ttu-id="f21c2-118">**Description**</span><span class="sxs-lookup"><span data-stu-id="f21c2-118">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="f21c2-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="f21c2-119">DefaultValue</span></span>|<span data-ttu-id="f21c2-120">URL</span><span class="sxs-lookup"><span data-stu-id="f21c2-120">URL</span></span>|<span data-ttu-id="f21c2-121">obligatoire</span><span class="sxs-lookup"><span data-stu-id="f21c2-121">required</span></span>|<span data-ttu-id="f21c2-122">Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="f21c2-122">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
