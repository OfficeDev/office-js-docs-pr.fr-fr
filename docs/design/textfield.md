---
title: Composant TextField dans Office UI Fabric
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 7c579bc12ed0cf1f4d4af52306c6556f7f79f427
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437289"
---
# <a name="textfield-component-in-office-ui-fabric"></a><span data-ttu-id="9499c-102">Composant TextField dans Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="9499c-102">TextField component in Office UI Fabric</span></span>

<span data-ttu-id="9499c-p101">Les champs de texte permettent aux utilisateurs de saisir du texte. Ils sont généralement utilisés pour capturer une seule ligne de texte mais peuvent être configurés pour capturer plusieurs lignes de texte. Le texte s’affiche à l’écran dans un format simple et uniforme.</span><span class="sxs-lookup"><span data-stu-id="9499c-p101">A text field enables users to type text. It's typically used to capture a single line of text but can be configured to capture multiple lines of text. The text displays on the screen in a simple, uniform format.</span></span>
  
#### <a name="example-textfield-in-a-task-pane"></a><span data-ttu-id="9499c-106">Exemple : TextField dans un volet Office</span><span class="sxs-lookup"><span data-stu-id="9499c-106">Example: TextField in a task pane</span></span>

![Image illustrant le composant TextField](../images/overview-with-app-text-field.png)

## <a name="best-practices"></a><span data-ttu-id="9499c-108">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="9499c-108">Best practices</span></span>

|<span data-ttu-id="9499c-109">**À faire**</span><span class="sxs-lookup"><span data-stu-id="9499c-109">**Do**</span></span>|<span data-ttu-id="9499c-110">**À ne pas faire**</span><span class="sxs-lookup"><span data-stu-id="9499c-110">**Don't**</span></span>|
|:------------|:--------------|
|<span data-ttu-id="9499c-111">Utiliser des champs de texte pour accepter la saisie de données sur un formulaire ou une page.</span><span class="sxs-lookup"><span data-stu-id="9499c-111">Use text fields to accept data input on a form or page.</span></span>|<span data-ttu-id="9499c-112">Ne pas utiliser de champs de texte pour rendre une copie de base dans un élément de corps d’une page.</span><span class="sxs-lookup"><span data-stu-id="9499c-112">Don’t use text fields to render basic copy as part of a body element of a page.</span></span>|
|<span data-ttu-id="9499c-113">Étiqueter les champs de texte avec des noms utiles.</span><span class="sxs-lookup"><span data-stu-id="9499c-113">Label text fields with helpful names.</span></span>|<span data-ttu-id="9499c-114">Ne pas utiliser de champs de texte pour saisir une date ou une heure.</span><span class="sxs-lookup"><span data-stu-id="9499c-114">Don’t use text fields for date or time entry.</span></span> <span data-ttu-id="9499c-115">Utiliser plutôt un sélecteur de date et heure.</span><span class="sxs-lookup"><span data-stu-id="9499c-115">Instead, use a date-time picker.</span></span>|
|<span data-ttu-id="9499c-116">Utiliser un texte de l’espace réservé concis pour spécifier le contenu qui doit être saisi.</span><span class="sxs-lookup"><span data-stu-id="9499c-116">Use concise placeholder text to specify what content should be entered.</span></span>|<span data-ttu-id="9499c-p103">Ne pas utiliser de champs de texte si des options d’entrée valides peuvent être prédéfinies. Utiliser plutôt une liste déroulante.</span><span class="sxs-lookup"><span data-stu-id="9499c-p103">Don’t use text fields if you can predefine valid input options. Instead, use a drop-down.</span></span>|
|<span data-ttu-id="9499c-119">Fournir tous les états appropriés pour les champs de texte (statique, pointage, focus, engagé, non disponible, erreur).</span><span class="sxs-lookup"><span data-stu-id="9499c-119">Provide all appropriate states for the text fields (static, hover, focus, engaged, unavailable, error).</span></span>||
|<span data-ttu-id="9499c-120">Marquer clairement les champs obligatoires et facultatifs.</span><span class="sxs-lookup"><span data-stu-id="9499c-120">Clearly mark required and optional text fields.</span></span>||
|<span data-ttu-id="9499c-p104">Si possible, mettre en forme les champs de texte en fonction du format de données attendu. Par exemple, lors de la capture d’un numéro de téléphone à 10 chiffres, utiliser trois champs distincts pour stocker les différentes parties du numéro de téléphone.</span><span class="sxs-lookup"><span data-stu-id="9499c-p104">Whenever possible, format text fields according to the expected data format. For example, when capturing a 10-digit phone number, use three separate fields to store the different parts of the phone number.</span></span>||

## <a name="variants"></a><span data-ttu-id="9499c-123">Variantes</span><span class="sxs-lookup"><span data-stu-id="9499c-123">Variants</span></span>

|<span data-ttu-id="9499c-124">**Variation**</span><span class="sxs-lookup"><span data-stu-id="9499c-124">**Variation**</span></span>|<span data-ttu-id="9499c-125">**Description**</span><span class="sxs-lookup"><span data-stu-id="9499c-125">**Description**</span></span>|<span data-ttu-id="9499c-126">**Exemple**</span><span class="sxs-lookup"><span data-stu-id="9499c-126">**Example**</span></span>|
|:------------|:--------------|:----------|
|<span data-ttu-id="9499c-127">**Default TextField (champ de texte par défaut)**</span><span class="sxs-lookup"><span data-stu-id="9499c-127">**Default TextField**</span></span>|<span data-ttu-id="9499c-128">À utiliser comme champ de texte par défaut.</span><span class="sxs-lookup"><span data-stu-id="9499c-128">Use as the default text field.</span></span>|![Image Default TextField (champ de texte par défaut)](../images/textfield-default.png)<br/>|
|<span data-ttu-id="9499c-130">**Disabled TextField (champ de texte désactivé)**</span><span class="sxs-lookup"><span data-stu-id="9499c-130">**Disabled TextField**</span></span>|<span data-ttu-id="9499c-131">À utiliser lorsque le champ de texte est désactivé.</span><span class="sxs-lookup"><span data-stu-id="9499c-131">Use when the text field is disabled.</span></span>|![Image Disabled TextField (champ de texte désactivé)](../images/textfield-disabled.png)<br/>|
|<span data-ttu-id="9499c-133">**Required TextField (champ de texte obligatoire)**</span><span class="sxs-lookup"><span data-stu-id="9499c-133">**Required TextField**</span></span>|<span data-ttu-id="9499c-134">À utiliser lorsque le champ de texte est obligatoire.</span><span class="sxs-lookup"><span data-stu-id="9499c-134">Use when the text field input is required.</span></span>|![Image Required TextField (champ de texte obligatoire)](../images/textfield-required.png)<br/>|
|<span data-ttu-id="9499c-136">**TextField with a placeholder (champ de texte avec un espace réservé)**</span><span class="sxs-lookup"><span data-stu-id="9499c-136">**TextField with a placeholder**</span></span>|<span data-ttu-id="9499c-137">À utiliser lorsqu’un texte de l’espace réservé est nécessaire.</span><span class="sxs-lookup"><span data-stu-id="9499c-137">Use when placeholder text is needed.</span></span>|![Image TextField with a placeholder (champ de texte avec un espace réservé)](../images/textfield-placeholder.png)<br/>|
|<span data-ttu-id="9499c-139">**TextField with multiple lines (champ de texte avec plusieurs lignes)**</span><span class="sxs-lookup"><span data-stu-id="9499c-139">**TextField with multiple lines**</span></span>|<span data-ttu-id="9499c-140">À utiliser lorsque plusieurs lignes de texte sont nécessaires.</span><span class="sxs-lookup"><span data-stu-id="9499c-140">Use when many lines of text are needed.</span></span>|![Image TextField with a placeholder (champ de texte avec un espace réservé)](../images/textfield-multi.png)<br/>|

## <a name="implementation"></a><span data-ttu-id="9499c-142">Implémentation</span><span class="sxs-lookup"><span data-stu-id="9499c-142">Implementation</span></span>

<span data-ttu-id="9499c-143">Pour plus d’informations, reportez-vous à [TextField](https://dev.office.com/fabric#/components/textfield) et [Démarrer avec un exemple de code Fabric React](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).</span><span class="sxs-lookup"><span data-stu-id="9499c-143">For details, see [TextField](https://dev.office.com/fabric#/components/textfield) and [Getting started with Fabric React code sample](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).</span></span>

## <a name="see-also"></a><span data-ttu-id="9499c-144">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="9499c-144">See also</span></span>

- [<span data-ttu-id="9499c-145">Modèles de conception UX</span><span class="sxs-lookup"><span data-stu-id="9499c-145">UX Design Patterns</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [<span data-ttu-id="9499c-146">Office UI Fabric dans des compléments Office</span><span class="sxs-lookup"><span data-stu-id="9499c-146">Office UI Fabric in Office Add-ins</span></span>](office-ui-fabric.md)
