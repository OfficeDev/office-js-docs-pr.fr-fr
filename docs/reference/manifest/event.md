---
title: Élément Event dans le fichier manifeste
description: Définit un gestionnaire d’événements dans un complément.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 02037a54ad4b7e91a3697b53b04fa30e8a4909a9
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718229"
---
# <a name="event-element"></a><span data-ttu-id="a82e9-103">Élément Event</span><span class="sxs-lookup"><span data-stu-id="a82e9-103">Event element</span></span>

<span data-ttu-id="a82e9-104">Définit un gestionnaire d’événements dans un complément.</span><span class="sxs-lookup"><span data-stu-id="a82e9-104">Defines an event handler in an add-in.</span></span>

> [!NOTE] 
> <span data-ttu-id="a82e9-105">L' `Event` élément est actuellement uniquement pris en charge par Outlook sur le Web dans Office 365.</span><span class="sxs-lookup"><span data-stu-id="a82e9-105">The `Event` element is currently only supported by Outlook on the web in Office 365.</span></span>

## <a name="attributes"></a><span data-ttu-id="a82e9-106">Attributs</span><span class="sxs-lookup"><span data-stu-id="a82e9-106">Attributes</span></span>

|  <span data-ttu-id="a82e9-107">Attribut</span><span class="sxs-lookup"><span data-stu-id="a82e9-107">Attribute</span></span>  |  <span data-ttu-id="a82e9-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="a82e9-108">Required</span></span>  |  <span data-ttu-id="a82e9-109">Description</span><span class="sxs-lookup"><span data-stu-id="a82e9-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="a82e9-110">Type</span><span class="sxs-lookup"><span data-stu-id="a82e9-110">Type</span></span>](#type-attribute)  |  <span data-ttu-id="a82e9-111">Oui</span><span class="sxs-lookup"><span data-stu-id="a82e9-111">Yes</span></span>  | <span data-ttu-id="a82e9-112">Indique l’événement à gérer.</span><span class="sxs-lookup"><span data-stu-id="a82e9-112">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="a82e9-113">FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="a82e9-113">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="a82e9-114">Oui</span><span class="sxs-lookup"><span data-stu-id="a82e9-114">Yes</span></span>  | <span data-ttu-id="a82e9-p101">Indique le style d’exécution du gestionnaire d’événements, asynchrone ou synchrone. Actuellement, seuls les gestionnaires d’événement synchrones sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="a82e9-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="a82e9-117">FunctionName</span><span class="sxs-lookup"><span data-stu-id="a82e9-117">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="a82e9-118">Oui</span><span class="sxs-lookup"><span data-stu-id="a82e9-118">Yes</span></span>  | <span data-ttu-id="a82e9-119">Indique le nom de la fonction du gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="a82e9-119">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="a82e9-120">Attribut Type</span><span class="sxs-lookup"><span data-stu-id="a82e9-120">Type attribute</span></span>

<span data-ttu-id="a82e9-p102">Obligatoire. Indique l’événement qui appelle le gestionnaire d’événements. Les valeurs possibles pour cet attribut sont répertoriées dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="a82e9-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="a82e9-124">Type d’événement</span><span class="sxs-lookup"><span data-stu-id="a82e9-124">Event type</span></span>  |  <span data-ttu-id="a82e9-125">Description</span><span class="sxs-lookup"><span data-stu-id="a82e9-125">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="a82e9-126">Le gestionnaire d’événements est appelé quand l’utilisateur envoie un message ou une convocation.</span><span class="sxs-lookup"><span data-stu-id="a82e9-126">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="a82e9-127">Attribut FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="a82e9-127">FunctionExecution attribute</span></span>

<span data-ttu-id="a82e9-128">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="a82e9-128">Required.</span></span> <span data-ttu-id="a82e9-129">DOIT être défini sur `synchronous`.</span><span class="sxs-lookup"><span data-stu-id="a82e9-129">MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="a82e9-130">Attribut FunctionName</span><span class="sxs-lookup"><span data-stu-id="a82e9-130">FunctionName attribute</span></span>

<span data-ttu-id="a82e9-p104">Obligatoire. Indique le nom de la fonction du gestionnaire d’événements. Cette valeur doit correspondre au nom d’une fonction dans le [fichier de fonction](functionfile.md) du complément.</span><span class="sxs-lookup"><span data-stu-id="a82e9-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```
