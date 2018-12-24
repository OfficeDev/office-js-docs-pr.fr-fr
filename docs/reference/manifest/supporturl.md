---
title: Élément SupportUrl dans le fichier manifest
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 00234ef9fe8960b9956e6a2595e2e2e71bfb97c6
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432668"
---
# <a name="supporturl-element"></a><span data-ttu-id="ac5e2-102">Élément SupportUrl</span><span class="sxs-lookup"><span data-stu-id="ac5e2-102">SupportUrl element</span></span>

<span data-ttu-id="ac5e2-103">Spécifie l’URL d’une page qui fournit des informations de prise en charge pour votre complément.</span><span class="sxs-lookup"><span data-stu-id="ac5e2-103">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="ac5e2-104">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="ac5e2-104">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="ac5e2-105">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="ac5e2-105">Contained in</span></span>

[<span data-ttu-id="ac5e2-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="ac5e2-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="ac5e2-107">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="ac5e2-107">Can contain</span></span>

|  <span data-ttu-id="ac5e2-108">Élément</span><span class="sxs-lookup"><span data-stu-id="ac5e2-108">Element</span></span> | <span data-ttu-id="ac5e2-109">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="ac5e2-109">Required</span></span> | <span data-ttu-id="ac5e2-110">Description</span><span class="sxs-lookup"><span data-stu-id="ac5e2-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ac5e2-111">Override</span><span class="sxs-lookup"><span data-stu-id="ac5e2-111">Override</span></span>](override.md)   | <span data-ttu-id="ac5e2-112">Non</span><span class="sxs-lookup"><span data-stu-id="ac5e2-112">No</span></span> | <span data-ttu-id="ac5e2-113">Spécifie le paramètre pour les URL de paramètres régionaux supplémentaires</span><span class="sxs-lookup"><span data-stu-id="ac5e2-113">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="ac5e2-114">Attributs</span><span class="sxs-lookup"><span data-stu-id="ac5e2-114">Attributes</span></span>

|<span data-ttu-id="ac5e2-115">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="ac5e2-115">**Attribute**</span></span>|<span data-ttu-id="ac5e2-116">**Type**</span><span class="sxs-lookup"><span data-stu-id="ac5e2-116">**Type**</span></span>|<span data-ttu-id="ac5e2-117">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="ac5e2-117">**Required**</span></span>|<span data-ttu-id="ac5e2-118">**Description**</span><span class="sxs-lookup"><span data-stu-id="ac5e2-118">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="ac5e2-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="ac5e2-119">DefaultValue</span></span>|<span data-ttu-id="ac5e2-120">URL</span><span class="sxs-lookup"><span data-stu-id="ac5e2-120">URL</span></span>|<span data-ttu-id="ac5e2-121">obligatoire</span><span class="sxs-lookup"><span data-stu-id="ac5e2-121">required</span></span>|<span data-ttu-id="ac5e2-122">Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="ac5e2-122">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
