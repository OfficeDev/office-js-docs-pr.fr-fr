---
title: Élément Event dans le fichier manifeste
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 51bbcd5a3d5abe60b850e88e4063e6bbc2da37bc
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450589"
---
# <a name="event-element"></a><span data-ttu-id="34f1e-102">Élément Event</span><span class="sxs-lookup"><span data-stu-id="34f1e-102">Event element</span></span>

<span data-ttu-id="34f1e-103">Définit un gestionnaire d’événements dans un complément.</span><span class="sxs-lookup"><span data-stu-id="34f1e-103">Defines an event handler in an add-in.</span></span>

> [!NOTE] 
> <span data-ttu-id="34f1e-104">L' `Event` élément est actuellement uniquement pris en charge par Outlook sur le Web dans Office 365.</span><span class="sxs-lookup"><span data-stu-id="34f1e-104">The `Event` element is currently only supported by Outlook on the web in Office 365.</span></span>

## <a name="attributes"></a><span data-ttu-id="34f1e-105">Attributs</span><span class="sxs-lookup"><span data-stu-id="34f1e-105">Attributes</span></span>

|  <span data-ttu-id="34f1e-106">Attribut</span><span class="sxs-lookup"><span data-stu-id="34f1e-106">Attribute</span></span>  |  <span data-ttu-id="34f1e-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="34f1e-107">Required</span></span>  |  <span data-ttu-id="34f1e-108">Description</span><span class="sxs-lookup"><span data-stu-id="34f1e-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="34f1e-109">Type</span><span class="sxs-lookup"><span data-stu-id="34f1e-109">Type</span></span>](#type-attribute)  |  <span data-ttu-id="34f1e-110">Oui</span><span class="sxs-lookup"><span data-stu-id="34f1e-110">Yes</span></span>  | <span data-ttu-id="34f1e-111">Indique l’événement à gérer.</span><span class="sxs-lookup"><span data-stu-id="34f1e-111">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="34f1e-112">FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="34f1e-112">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="34f1e-113">Oui</span><span class="sxs-lookup"><span data-stu-id="34f1e-113">Yes</span></span>  | <span data-ttu-id="34f1e-p101">Indique le style d’exécution du gestionnaire d’événements, asynchrone ou synchrone. Actuellement, seuls les gestionnaires d’événement synchrones sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="34f1e-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="34f1e-116">FunctionName</span><span class="sxs-lookup"><span data-stu-id="34f1e-116">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="34f1e-117">Oui</span><span class="sxs-lookup"><span data-stu-id="34f1e-117">Yes</span></span>  | <span data-ttu-id="34f1e-118">Indique le nom de la fonction du gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="34f1e-118">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="34f1e-119">Attribut Type</span><span class="sxs-lookup"><span data-stu-id="34f1e-119">Type attribute</span></span>

<span data-ttu-id="34f1e-p102">Obligatoire. Indique l’événement qui appelle le gestionnaire d’événements. Les valeurs possibles pour cet attribut sont répertoriées dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="34f1e-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="34f1e-123">Type d’événement</span><span class="sxs-lookup"><span data-stu-id="34f1e-123">Event type</span></span>  |  <span data-ttu-id="34f1e-124">Description</span><span class="sxs-lookup"><span data-stu-id="34f1e-124">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="34f1e-125">Le gestionnaire d’événements est appelé quand l’utilisateur envoie un message ou une convocation.</span><span class="sxs-lookup"><span data-stu-id="34f1e-125">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="34f1e-126">Attribut FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="34f1e-126">FunctionExecution attribute</span></span>

<span data-ttu-id="34f1e-127">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="34f1e-127">Required.</span></span> <span data-ttu-id="34f1e-128">DOIT être défini sur `synchronous`.</span><span class="sxs-lookup"><span data-stu-id="34f1e-128">MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="34f1e-129">Attribut FunctionName</span><span class="sxs-lookup"><span data-stu-id="34f1e-129">FunctionName attribute</span></span>

<span data-ttu-id="34f1e-p104">Obligatoire. Indique le nom de la fonction du gestionnaire d’événements. Cette valeur doit correspondre au nom d’une fonction dans le [fichier de fonction](functionfile.md) du complément.</span><span class="sxs-lookup"><span data-stu-id="34f1e-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```
