---
title: Élément Event dans le fichier manifeste
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: eda895b01e106d67eef70f199be64086e9372bef
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432738"
---
# <a name="event-element"></a><span data-ttu-id="8a625-102">Élément Event</span><span class="sxs-lookup"><span data-stu-id="8a625-102">Event element</span></span>

<span data-ttu-id="8a625-103">Définit un gestionnaire d’événements dans un complément.</span><span class="sxs-lookup"><span data-stu-id="8a625-103">Defines an event handler in an add-in.</span></span>

> [!NOTE] 
> <span data-ttu-id="8a625-104">L’élément `Event` est actuellement uniquement pris en charge par Outlook sur le web dans Office 365.</span><span class="sxs-lookup"><span data-stu-id="8a625-104">Note: The `Event` element is currently only supported by Outlook on the web in Office 365.</span></span>

## <a name="attributes"></a><span data-ttu-id="8a625-105">Attributs</span><span class="sxs-lookup"><span data-stu-id="8a625-105">Attributes</span></span>

|  <span data-ttu-id="8a625-106">Attribut</span><span class="sxs-lookup"><span data-stu-id="8a625-106">Attribute</span></span>  |  <span data-ttu-id="8a625-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="8a625-107">Required</span></span>  |  <span data-ttu-id="8a625-108">Description</span><span class="sxs-lookup"><span data-stu-id="8a625-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="8a625-109">Type</span><span class="sxs-lookup"><span data-stu-id="8a625-109">Type</span></span>](#type-attribute)  |  <span data-ttu-id="8a625-110">Oui</span><span class="sxs-lookup"><span data-stu-id="8a625-110">Yes</span></span>  | <span data-ttu-id="8a625-111">Indique l’événement à gérer.</span><span class="sxs-lookup"><span data-stu-id="8a625-111">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="8a625-112">FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="8a625-112">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="8a625-113">Oui</span><span class="sxs-lookup"><span data-stu-id="8a625-113">Yes</span></span>  | <span data-ttu-id="8a625-p101">Indique le style d’exécution du gestionnaire d’événements, asynchrone ou synchrone. Actuellement, seuls les gestionnaires d’événement synchrones sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="8a625-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="8a625-116">FunctionName</span><span class="sxs-lookup"><span data-stu-id="8a625-116">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="8a625-117">Oui</span><span class="sxs-lookup"><span data-stu-id="8a625-117">Yes</span></span>  | <span data-ttu-id="8a625-118">Indique le nom de la fonction du gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="8a625-118">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="8a625-119">Attribut Type</span><span class="sxs-lookup"><span data-stu-id="8a625-119">Type attribute</span></span>

<span data-ttu-id="8a625-p102">Obligatoire. Indique l’événement qui appelle le gestionnaire d’événements. Les valeurs possibles pour cet attribut sont répertoriées dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="8a625-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="8a625-123">Type d’événement</span><span class="sxs-lookup"><span data-stu-id="8a625-123">Event type</span></span>  |  <span data-ttu-id="8a625-124">Description</span><span class="sxs-lookup"><span data-stu-id="8a625-124">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="8a625-125">Le gestionnaire d’événements est appelé quand l’utilisateur envoie un message ou une convocation.</span><span class="sxs-lookup"><span data-stu-id="8a625-125">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="8a625-126">Attribut FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="8a625-126">FunctionExecution attribute</span></span>

<span data-ttu-id="8a625-p103">Obligatoire. DOIT être défini sur `synchronous`.</span><span class="sxs-lookup"><span data-stu-id="8a625-p103">Required. MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="8a625-129">Attribut FunctionName</span><span class="sxs-lookup"><span data-stu-id="8a625-129">FunctionName attribute</span></span>

<span data-ttu-id="8a625-p104">Obligatoire. Indique le nom de la fonction du gestionnaire d’événements. Cette valeur doit correspondre au nom d’une fonction dans le [fichier de fonction](functionfile.md) du complément.</span><span class="sxs-lookup"><span data-stu-id="8a625-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```