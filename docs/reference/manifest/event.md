---
title: Élément Event dans le fichier manifeste
description: Définit un gestionnaire d’événements dans un complément.
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: 80f21d1819e3d7e335389070ccac0db583026045
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275706"
---
# <a name="event-element"></a><span data-ttu-id="e6b5d-103">Élément Event</span><span class="sxs-lookup"><span data-stu-id="e6b5d-103">Event element</span></span>

<span data-ttu-id="e6b5d-104">Définit un gestionnaire d’événements dans un complément.</span><span class="sxs-lookup"><span data-stu-id="e6b5d-104">Defines an event handler in an add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="e6b5d-105">Pour plus d’informations sur la prise en charge et l’utilisation, consultez la rubrique relative à la [fonctionnalité d’envoi pour les compléments Outlook](../../outlook/outlook-on-send-addins.md).</span><span class="sxs-lookup"><span data-stu-id="e6b5d-105">For information about support and usage, see [On-send feature for Outlook add-ins](../../outlook/outlook-on-send-addins.md).</span></span>

## <a name="attributes"></a><span data-ttu-id="e6b5d-106">Attributs</span><span class="sxs-lookup"><span data-stu-id="e6b5d-106">Attributes</span></span>

|  <span data-ttu-id="e6b5d-107">Attribut</span><span class="sxs-lookup"><span data-stu-id="e6b5d-107">Attribute</span></span>  |  <span data-ttu-id="e6b5d-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="e6b5d-108">Required</span></span>  |  <span data-ttu-id="e6b5d-109">Description</span><span class="sxs-lookup"><span data-stu-id="e6b5d-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="e6b5d-110">Type</span><span class="sxs-lookup"><span data-stu-id="e6b5d-110">Type</span></span>](#type-attribute)  |  <span data-ttu-id="e6b5d-111">Oui</span><span class="sxs-lookup"><span data-stu-id="e6b5d-111">Yes</span></span>  | <span data-ttu-id="e6b5d-112">Indique l’événement à gérer.</span><span class="sxs-lookup"><span data-stu-id="e6b5d-112">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="e6b5d-113">FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="e6b5d-113">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="e6b5d-114">Oui</span><span class="sxs-lookup"><span data-stu-id="e6b5d-114">Yes</span></span>  | <span data-ttu-id="e6b5d-p101">Indique le style d’exécution du gestionnaire d’événements, asynchrone ou synchrone. Actuellement, seuls les gestionnaires d’événement synchrones sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="e6b5d-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="e6b5d-117">FunctionName</span><span class="sxs-lookup"><span data-stu-id="e6b5d-117">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="e6b5d-118">Oui</span><span class="sxs-lookup"><span data-stu-id="e6b5d-118">Yes</span></span>  | <span data-ttu-id="e6b5d-119">Indique le nom de la fonction du gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="e6b5d-119">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="e6b5d-120">Attribut Type</span><span class="sxs-lookup"><span data-stu-id="e6b5d-120">Type attribute</span></span>

<span data-ttu-id="e6b5d-p102">Obligatoire. Indique l’événement qui appelle le gestionnaire d’événements. Les valeurs possibles pour cet attribut sont répertoriées dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="e6b5d-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="e6b5d-124">Type d’événement</span><span class="sxs-lookup"><span data-stu-id="e6b5d-124">Event type</span></span>  |  <span data-ttu-id="e6b5d-125">Description</span><span class="sxs-lookup"><span data-stu-id="e6b5d-125">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="e6b5d-126">Le gestionnaire d’événements est appelé quand l’utilisateur envoie un message ou une convocation.</span><span class="sxs-lookup"><span data-stu-id="e6b5d-126">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="e6b5d-127">Attribut FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="e6b5d-127">FunctionExecution attribute</span></span>

<span data-ttu-id="e6b5d-128">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="e6b5d-128">Required.</span></span> <span data-ttu-id="e6b5d-129">DOIT être défini sur `synchronous`.</span><span class="sxs-lookup"><span data-stu-id="e6b5d-129">MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="e6b5d-130">Attribut FunctionName</span><span class="sxs-lookup"><span data-stu-id="e6b5d-130">FunctionName attribute</span></span>

<span data-ttu-id="e6b5d-p104">Obligatoire. Indique le nom de la fonction du gestionnaire d’événements. Cette valeur doit correspondre au nom d’une fonction dans le [fichier de fonction](functionfile.md) du complément.</span><span class="sxs-lookup"><span data-stu-id="e6b5d-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
```
