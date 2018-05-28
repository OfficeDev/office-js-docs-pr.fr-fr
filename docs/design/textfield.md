---
title: Composant TextField dans Office UI Fabric
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 7c579bc12ed0cf1f4d4af52306c6556f7f79f427
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="textfield-component-in-office-ui-fabric"></a><span data-ttu-id="91ae4-102">Composant TextField dans Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="91ae4-102">TextField component in Office UI Fabric</span></span>

<span data-ttu-id="91ae4-p101">Les champs de texte permettent aux utilisateurs de saisir du texte. Ils sont g?n?ralement utilis?s pour capturer une seule ligne de texte mais peuvent ?tre configur?s pour capturer plusieurs lignes de texte. Le texte s?affiche ? l??cran dans un format simple et uniforme.</span><span class="sxs-lookup"><span data-stu-id="91ae4-p101">A text field enables users to type text. It's typically used to capture a single line of text but can be configured to capture multiple lines of text. The text displays on the screen in a simple, uniform format.</span></span>
  
#### <a name="example-textfield-in-a-task-pane"></a><span data-ttu-id="91ae4-106">Exemple : TextField dans un volet Office</span><span class="sxs-lookup"><span data-stu-id="91ae4-106">Example: TextField in a task pane</span></span>

![Image illustrant le composant TextField](../images/overview-with-app-text-field.png)

## <a name="best-practices"></a><span data-ttu-id="91ae4-108">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="91ae4-108">Best practices</span></span>

|<span data-ttu-id="91ae4-109">**? faire**</span><span class="sxs-lookup"><span data-stu-id="91ae4-109">**Do**</span></span>|<span data-ttu-id="91ae4-110">**? ne pas faire**</span><span class="sxs-lookup"><span data-stu-id="91ae4-110">**Don't**</span></span>|
|:------------|:--------------|
|<span data-ttu-id="91ae4-111">Utiliser des champs de texte pour accepter la saisie de donn?es sur un formulaire ou une page.</span><span class="sxs-lookup"><span data-stu-id="91ae4-111">Use text fields to accept data input on a form or page.</span></span>|<span data-ttu-id="91ae4-112">Ne pas utiliser de champs de texte pour rendre une copie de base dans un ?l?ment de corps d?une page.</span><span class="sxs-lookup"><span data-stu-id="91ae4-112">Don?t use text fields to render basic copy as part of a body element of a page.</span></span>|
|<span data-ttu-id="91ae4-113">?tiqueter les champs de texte avec des noms utiles.</span><span class="sxs-lookup"><span data-stu-id="91ae4-113">Label text fields with helpful names.</span></span>|<span data-ttu-id="91ae4-114">Ne pas utiliser de champs de texte pour saisir une date ou une heure.</span><span class="sxs-lookup"><span data-stu-id="91ae4-114">Don?t use text fields for date or time entry.</span></span> <span data-ttu-id="91ae4-115">Utiliser plut?t un s?lecteur de date et heure.</span><span class="sxs-lookup"><span data-stu-id="91ae4-115">Instead, use a date-time picker.</span></span>|
|<span data-ttu-id="91ae4-116">Utiliser un texte de l?espace r?serv? concis pour sp?cifier le contenu qui doit ?tre saisi.</span><span class="sxs-lookup"><span data-stu-id="91ae4-116">Use concise placeholder text to specify what content should be entered.</span></span>|<span data-ttu-id="91ae4-p103">Ne pas utiliser de champs de texte si des options d?entr?e valides peuvent ?tre pr?d?finies. Utiliser plut?t une liste d?roulante.</span><span class="sxs-lookup"><span data-stu-id="91ae4-p103">Don?t use text fields if you can predefine valid input options. Instead, use a drop-down.</span></span>|
|<span data-ttu-id="91ae4-119">Fournir tous les ?tats appropri?s pour les champs de texte (statique, pointage, focus, engag?, non disponible, erreur).</span><span class="sxs-lookup"><span data-stu-id="91ae4-119">Provide all appropriate states for the text fields (static, hover, focus, engaged, unavailable, error).</span></span>||
|<span data-ttu-id="91ae4-120">Marquer clairement les champs obligatoires et facultatifs.</span><span class="sxs-lookup"><span data-stu-id="91ae4-120">Clearly mark required and optional text fields.</span></span>||
|<span data-ttu-id="91ae4-p104">Si possible, mettre en forme les champs de texte en fonction du format de donn?es attendu. Par exemple, lors de la capture d?un num?ro de t?l?phone ? 10 chiffres, utiliser trois champs distincts pour stocker les diff?rentes parties du num?ro de t?l?phone.</span><span class="sxs-lookup"><span data-stu-id="91ae4-p104">Whenever possible, format text fields according to the expected data format. For example, when capturing a 10-digit phone number, use three separate fields to store the different parts of the phone number.</span></span>||

## <a name="variants"></a><span data-ttu-id="91ae4-123">Variantes</span><span class="sxs-lookup"><span data-stu-id="91ae4-123">Variants</span></span>

|<span data-ttu-id="91ae4-124">**Variation**</span><span class="sxs-lookup"><span data-stu-id="91ae4-124">**Variation**</span></span>|<span data-ttu-id="91ae4-125">**Description**</span><span class="sxs-lookup"><span data-stu-id="91ae4-125">**Description**</span></span>|<span data-ttu-id="91ae4-126">**Exemple**</span><span class="sxs-lookup"><span data-stu-id="91ae4-126">**Example**</span></span>|
|:------------|:--------------|:----------|
|<span data-ttu-id="91ae4-127">**Default TextField (champ de texte par d?faut)**</span><span class="sxs-lookup"><span data-stu-id="91ae4-127">**Default TextField**</span></span>|<span data-ttu-id="91ae4-128">? utiliser comme champ de texte par d?faut.</span><span class="sxs-lookup"><span data-stu-id="91ae4-128">Use as the default text field.</span></span>|![Image Default TextField (champ de texte par d?faut)](../images/textfield-default.png)<br/>|
|<span data-ttu-id="91ae4-130">**Disabled TextField (champ de texte d?sactiv?)**</span><span class="sxs-lookup"><span data-stu-id="91ae4-130">**Disabled TextField**</span></span>|<span data-ttu-id="91ae4-131">? utiliser lorsque le champ de texte est d?sactiv?.</span><span class="sxs-lookup"><span data-stu-id="91ae4-131">Use when the text field is disabled.</span></span>|![Image Disabled TextField (champ de texte d?sactiv?)](../images/textfield-disabled.png)<br/>|
|<span data-ttu-id="91ae4-133">**Required TextField (champ de texte obligatoire)**</span><span class="sxs-lookup"><span data-stu-id="91ae4-133">**Required TextField**</span></span>|<span data-ttu-id="91ae4-134">? utiliser lorsque le champ de texte est obligatoire.</span><span class="sxs-lookup"><span data-stu-id="91ae4-134">Use when the text field input is required.</span></span>|![Image Required TextField (champ de texte obligatoire)](../images/textfield-required.png)<br/>|
|<span data-ttu-id="91ae4-136">**TextField with a placeholder (champ de texte avec un espace r?serv?)**</span><span class="sxs-lookup"><span data-stu-id="91ae4-136">**TextField with a placeholder**</span></span>|<span data-ttu-id="91ae4-137">? utiliser lorsqu?un texte de l?espace r?serv? est n?cessaire.</span><span class="sxs-lookup"><span data-stu-id="91ae4-137">Use when placeholder text is needed.</span></span>|![Image TextField with a placeholder (champ de texte avec un espace r?serv?)](../images/textfield-placeholder.png)<br/>|
|<span data-ttu-id="91ae4-139">**TextField with multiple lines (champ de texte avec plusieurs lignes)**</span><span class="sxs-lookup"><span data-stu-id="91ae4-139">**TextField with multiple lines**</span></span>|<span data-ttu-id="91ae4-140">? utiliser lorsque plusieurs lignes de texte sont n?cessaires.</span><span class="sxs-lookup"><span data-stu-id="91ae4-140">Use when many lines of text are needed.</span></span>|![Image TextField with a placeholder (champ de texte avec un espace r?serv?)](../images/textfield-multi.png)<br/>|

## <a name="implementation"></a><span data-ttu-id="91ae4-142">Impl?mentation</span><span class="sxs-lookup"><span data-stu-id="91ae4-142">Implementation</span></span>

<span data-ttu-id="91ae4-143">Pour plus d?informations, reportez-vous ? [TextField](https://dev.office.com/fabric#/components/textfield) et [D?marrer avec un exemple de code Fabric React](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).</span><span class="sxs-lookup"><span data-stu-id="91ae4-143">For details, see [TextField](https://dev.office.com/fabric#/components/textfield) and [Getting started with Fabric React code sample](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).</span></span>

## <a name="see-also"></a><span data-ttu-id="91ae4-144">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="91ae4-144">See also</span></span>

- [<span data-ttu-id="91ae4-145">Mod?les de conception UX</span><span class="sxs-lookup"><span data-stu-id="91ae4-145">UX Design Patterns</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [<span data-ttu-id="91ae4-146">Office UI Fabric dans des compl?ments Office</span><span class="sxs-lookup"><span data-stu-id="91ae4-146">Office UI Fabric in Office Add-ins</span></span>](office-ui-fabric.md)
