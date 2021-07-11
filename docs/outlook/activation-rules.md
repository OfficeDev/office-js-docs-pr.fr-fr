---
title: Règles d’activation pour les compléments Outlook
description: Outlook active certains types de complément si le message ou le rendez-vous que l’utilisateur lit ou compose respecte les règles d’activation du complément.
ms.date: 09/22/2020
localization_priority: Normal
ms.openlocfilehash: 24f17b7bb3da4665f3f05b23d34ba15bcc4ae729
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349020"
---
# <a name="activation-rules-for-contextual-outlook-add-ins"></a><span data-ttu-id="df5fc-103">Règles d’activation des compléments contextuels Outlook </span><span class="sxs-lookup"><span data-stu-id="df5fc-103">Activation rules for contextual Outlook add-ins</span></span>

<span data-ttu-id="df5fc-p101">Outlook active certains types de compléments si le message ou le rendez-vous que l’utilisateur lit ou compose respecte les règles d’activation du complément. Cela est vrai pour tous les compléments qui utilisent le schéma de manifeste 1.1. L’utilisateur peut choisir le complément à partir de l’interface utilisateur Outlook afin de le démarrer pour l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="df5fc-p101">Outlook activates some types of add-ins if the message or appointment that the user is reading or composing satisfies the activation rules of the add-in. This is true for all add-ins that use the 1.1 manifest schema. The user can then choose the add-in from the Outlook UI to start it for the current item.</span></span>

<span data-ttu-id="df5fc-107">La figure suivante illustre les compléments Outlook activés dans la barre des compléments pour le message dans le volet de lecture.</span><span class="sxs-lookup"><span data-stu-id="df5fc-107">The following figure shows Outlook add-ins activated in the add-in bar for the message in the Reading Pane.</span></span>

![Barre de l’application affichant les applications de messagerie en lecture activées.](../images/read-form-app-bar.png)


## <a name="specify-activation-rules-in-a-manifest"></a><span data-ttu-id="df5fc-109">Spécifier des règles d’activation dans un manifeste</span><span class="sxs-lookup"><span data-stu-id="df5fc-109">Specify activation rules in a manifest</span></span>


<span data-ttu-id="df5fc-110">Pour que Outlook activer un complément pour des conditions spécifiques, spécifiez des règles d’activation dans le manifeste du complément à l’aide de l’un des éléments `Rule` suivants.</span><span class="sxs-lookup"><span data-stu-id="df5fc-110">To have Outlook activate an add-in for specific conditions, specify activation rules in the add-in manifest by using one of the following `Rule` elements.</span></span>

- <span data-ttu-id="df5fc-111">[Élément de règle (MailApp complexType)](../reference/manifest/rule.md) : spécifie une règle individuelle.</span><span class="sxs-lookup"><span data-stu-id="df5fc-111">[Rule element (MailApp complexType)](../reference/manifest/rule.md) - Specifies an individual rule.</span></span>
- <span data-ttu-id="df5fc-112">[Élément de règle (RuleCollection complexType)](../reference/manifest/rule.md#rulecollection) : combine plusieurs règles à l’aide d’opérations logiques.</span><span class="sxs-lookup"><span data-stu-id="df5fc-112">[Rule element (RuleCollection complexType)](../reference/manifest/rule.md#rulecollection) - Combines multiple rules using logical operations.</span></span>


 > [!NOTE]
 > <span data-ttu-id="df5fc-113">`Rule`L’élément que vous utilisez pour spécifier une règle individuelle est du type complexe [Rule](../reference/manifest/rule.md) abstrait.</span><span class="sxs-lookup"><span data-stu-id="df5fc-113">The `Rule` element that you use to specify an individual rule is of the abstract [Rule](../reference/manifest/rule.md) complex type.</span></span> <span data-ttu-id="df5fc-114">Chacun des types de règles suivants étend ce `Rule` type complexe abstrait.</span><span class="sxs-lookup"><span data-stu-id="df5fc-114">Each of the following types of rules extends this abstract `Rule` complex type.</span></span> <span data-ttu-id="df5fc-115">Ainsi, quand vous spécifiez une règle individuelle dans un manifeste, vous devez utiliser l’attribut [xsi:type](https://www.w3.org/TR/xmlschema-1/) pour définir plus précisément l’un des types de règle suivants.</span><span class="sxs-lookup"><span data-stu-id="df5fc-115">So when you specify an individual rule in a manifest, you must use the [xsi:type](https://www.w3.org/TR/xmlschema-1/) attribute to further define one of the following types of rules.</span></span>
 > 
 > <span data-ttu-id="df5fc-116">Par exemple, la règle suivante définit une [règle ItemIs.](../reference/manifest/rule.md#itemis-rule)</span><span class="sxs-lookup"><span data-stu-id="df5fc-116">For example, the following rule defines an [ItemIs](../reference/manifest/rule.md#itemis-rule) rule.</span></span>
 > `<Rule xsi:type="ItemIs" ItemType="Message" />`
 > 
 > <span data-ttu-id="df5fc-117">`FormType`L’attribut s’applique aux règles d’activation dans le manifeste v1.1, mais n’est pas défini dans `VersionOverrides` la v1.0.</span><span class="sxs-lookup"><span data-stu-id="df5fc-117">The `FormType` attribute applies to activation rules in the manifest v1.1 but is not defined in `VersionOverrides` v1.0.</span></span> <span data-ttu-id="df5fc-118">Il ne peut donc pas être utilisé lorsque [itemIs](../reference/manifest/rule.md#itemis-rule) est utilisé dans `VersionOverrides` le nœud.</span><span class="sxs-lookup"><span data-stu-id="df5fc-118">So it can't be used when [ItemIs](../reference/manifest/rule.md#itemis-rule) is used in the `VersionOverrides` node.</span></span>

<span data-ttu-id="df5fc-p105">Le tableau suivant répertorie les types de règle disponibles. Vous trouverez plus d’informations dans le tableau et dans les articles indiqués sous [Créer des compléments Outlook pour des formulaires de lecture](read-scenario.md).</span><span class="sxs-lookup"><span data-stu-id="df5fc-p105">The following table lists the types of rules that are available. You can find more information following the table and in the specified articles under [Create Outlook add-ins for read forms](read-scenario.md).</span></span>

<br/>

|<span data-ttu-id="df5fc-121">**Nom de la règle**</span><span class="sxs-lookup"><span data-stu-id="df5fc-121">**Rule name**</span></span>|<span data-ttu-id="df5fc-122">**Formulaires applicables**</span><span class="sxs-lookup"><span data-stu-id="df5fc-122">**Applicable forms**</span></span>|<span data-ttu-id="df5fc-123">**Description**</span><span class="sxs-lookup"><span data-stu-id="df5fc-123">**Description**</span></span>|
|:-----|:-----|:-----|
|[<span data-ttu-id="df5fc-124">ItemIs</span><span class="sxs-lookup"><span data-stu-id="df5fc-124">ItemIs</span></span>](#itemis-rule)|<span data-ttu-id="df5fc-125">Lecture, composition</span><span class="sxs-lookup"><span data-stu-id="df5fc-125">Read, Compose</span></span>|<span data-ttu-id="df5fc-p106">Vérifie si l’élément actuel est du type spécifié (message ou rendez-vous). Peut également vérifier la classe de l’élément et le type de formulaire, ainsi qu’éventuellement la classe de message de l’élément.</span><span class="sxs-lookup"><span data-stu-id="df5fc-p106">Checks to see whether the current item is of the specified type (message or appointment). Can also check the item class and form type.and optionally, item message class.</span></span>|
|[<span data-ttu-id="df5fc-128">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="df5fc-128">ItemHasAttachment</span></span>](#itemhasattachment-rule)|<span data-ttu-id="df5fc-129">Lecture</span><span class="sxs-lookup"><span data-stu-id="df5fc-129">Read</span></span>|<span data-ttu-id="df5fc-130">Vérifie si l’élément sélectionné contient une pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="df5fc-130">Checks to see whether the selected item contains an attachment.</span></span>|
|[<span data-ttu-id="df5fc-131">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="df5fc-131">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)|<span data-ttu-id="df5fc-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="df5fc-132">Read</span></span>|<span data-ttu-id="df5fc-p107">Vérifie si l’élément sélectionné contient une ou plusieurs entités reconnues. Plus d’informations : [Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues](match-strings-in-an-item-as-well-known-entities.md).</span><span class="sxs-lookup"><span data-stu-id="df5fc-p107">Checks to see whether the selected item contains one or more well-known entities. More information: [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md).</span></span>|
|[<span data-ttu-id="df5fc-135">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="df5fc-135">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)|<span data-ttu-id="df5fc-136">Lecture</span><span class="sxs-lookup"><span data-stu-id="df5fc-136">Read</span></span>|<span data-ttu-id="df5fc-137">Vérifie si l’adresse électronique de l’expéditeur, l’objet et/ou le corps de l’élément sélectionné contient une correspondance avec une expression régulière.Plus d’informations : [Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook](use-regular-expressions-to-show-an-outlook-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="df5fc-137">Checks to see whether the sender's email address, the subject, and/or the body of the selected item contains a match to a regular expression.More information: [Use regular expression activation rules to show an Outlook add-in](use-regular-expressions-to-show-an-outlook-add-in.md).</span></span>|
|[<span data-ttu-id="df5fc-138">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="df5fc-138">RuleCollection</span></span>](#rulecollection-rule)|<span data-ttu-id="df5fc-139">Lecture, composition</span><span class="sxs-lookup"><span data-stu-id="df5fc-139">Read, Compose</span></span>|<span data-ttu-id="df5fc-140">Associe un ensemble de règles pour vous permettre de former des règles plus complexes.</span><span class="sxs-lookup"><span data-stu-id="df5fc-140">Combines a set of rules so that you can form more complex rules.</span></span>|

## <a name="itemis-rule"></a><span data-ttu-id="df5fc-141">Règle ItemIs</span><span class="sxs-lookup"><span data-stu-id="df5fc-141">ItemIs rule</span></span>

<span data-ttu-id="df5fc-142">Le type complexe **ItemIs** définit une règle qui a pour valeur **true** si l’élément actuel correspond au type d’élément et, éventuellement, la classe de message de l’élément s’il est indiqué dans la règle.</span><span class="sxs-lookup"><span data-stu-id="df5fc-142">The **ItemIs** complex type defines a rule that evaluates to **true** if the current item matches the item type, and optionally the item message class if it's stated in the rule.</span></span>

<span data-ttu-id="df5fc-143">Spécifiez l’un des types d’éléments suivants dans `ItemType` l’attribut d’une **règle ItemIs.**</span><span class="sxs-lookup"><span data-stu-id="df5fc-143">Specify one of the following item types in the `ItemType` attribute of an **ItemIs** rule.</span></span> <span data-ttu-id="df5fc-144">Vous pouvez spécifier plusieurs règles **ItemIs** dans un manifeste.</span><span class="sxs-lookup"><span data-stu-id="df5fc-144">You can specify more than one **ItemIs** rule in a manifest.</span></span> <span data-ttu-id="df5fc-145">L’élément ItemType simpleType définit les types d’élément Outlook qui prennent en charge les compléments Outlook.</span><span class="sxs-lookup"><span data-stu-id="df5fc-145">The ItemType simpleType defines the types of Outlook items that support Outlook add-ins.</span></span>

<br/>

|<span data-ttu-id="df5fc-146">**Value**</span><span class="sxs-lookup"><span data-stu-id="df5fc-146">**Value**</span></span>|<span data-ttu-id="df5fc-147">**Description**</span><span class="sxs-lookup"><span data-stu-id="df5fc-147">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="df5fc-148">**Rendez-vous**</span><span class="sxs-lookup"><span data-stu-id="df5fc-148">**Appointment**</span></span>|<span data-ttu-id="df5fc-149">Spécifie un élément dans le calendrier Outlook.</span><span class="sxs-lookup"><span data-stu-id="df5fc-149">Specifies an item in an Outlook calendar.</span></span> <span data-ttu-id="df5fc-150">Par exemple, un élément de réunion auquel une réponse a été donnée et auquel un organisateur et des participants sont associés, ou un rendez-vous auquel n’est associé aucun organisateur ou participant et qui constitue un simple élément de calendrier.</span><span class="sxs-lookup"><span data-stu-id="df5fc-150">This includes a meeting item that has been responded to and has an organizer and attendees, or an appointment that does not have an organizer or attendee and is simply an item on the calendar.</span></span> <span data-ttu-id="df5fc-151">Cela correspond à la classe de message IPM.Appointment dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="df5fc-151">This corresponds to the IPM.Appointment message class in Outlook.</span></span>|
|<span data-ttu-id="df5fc-152">**Message**</span><span class="sxs-lookup"><span data-stu-id="df5fc-152">**Message**</span></span>|<span data-ttu-id="df5fc-153">Spécifie l’un des éléments suivants reçus dans la boîte de réception.</span><span class="sxs-lookup"><span data-stu-id="df5fc-153">Specifies one of the following items received in typically the Inbox.</span></span> <ul><li><p><span data-ttu-id="df5fc-p110">Message électronique. Cela correspond à la classe de message IPM.Note dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="df5fc-p110">An email message. This corresponds to the IPM.Note message class in Outlook.</span></span></p></li><li><p><span data-ttu-id="df5fc-156">Demande de réunion, réponse à une demande de réunion ou annulation d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="df5fc-156">A meeting request, response, or cancellation.</span></span> <span data-ttu-id="df5fc-157">Cela correspond aux classes de message suivantes dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="df5fc-157">This corresponds to the following message classes in Outlook.</span></span></p><p><span data-ttu-id="df5fc-158">IPM.Schedule.Meeting.Request</span><span class="sxs-lookup"><span data-stu-id="df5fc-158">IPM.Schedule.Meeting.Request</span></span></p><p><span data-ttu-id="df5fc-159">IPM.Schedule.Meeting.Neg</span><span class="sxs-lookup"><span data-stu-id="df5fc-159">IPM.Schedule.Meeting.Neg</span></span></p><p><span data-ttu-id="df5fc-160">IPM.Schedule.Meeting.Pos</span><span class="sxs-lookup"><span data-stu-id="df5fc-160">IPM.Schedule.Meeting.Pos</span></span></p><p><span data-ttu-id="df5fc-161">IPM.Schedule.Meeting.Tent</span><span class="sxs-lookup"><span data-stu-id="df5fc-161">IPM.Schedule.Meeting.Tent</span></span></p><p><span data-ttu-id="df5fc-162">IPM.Schedule.Meeting.Canceled</span><span class="sxs-lookup"><span data-stu-id="df5fc-162">IPM.Schedule.Meeting.Canceled</span></span></p></li></ul>|

<span data-ttu-id="df5fc-163">L’attribut permet de spécifier le mode (lecture ou composition) dans lequel le `FormType` module doit être activé.</span><span class="sxs-lookup"><span data-stu-id="df5fc-163">The `FormType` attribute is used to specify the mode (read or compose) in which the add-in should activate.</span></span>


 > [!NOTE]
 > <span data-ttu-id="df5fc-164">L’attribut ItemIs `FormType` est défini dans les schémas v1.1 et ultérieur, mais pas dans la `VersionOverrides` v1.0.</span><span class="sxs-lookup"><span data-stu-id="df5fc-164">The ItemIs `FormType` attribute is defined in schema v1.1 and later but not in `VersionOverrides` v1.0.</span></span> <span data-ttu-id="df5fc-165">N’incluez pas `FormType` l’attribut lors de la définition des commandes de add-in.</span><span class="sxs-lookup"><span data-stu-id="df5fc-165">Do not include the `FormType` attribute when defining add-in commands.</span></span>

<span data-ttu-id="df5fc-166">Une fois qu’un complément est activé, vous pouvez utiliser la propriété [mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) pour obtenir l’élément actuellement sélectionné dans Outlook, et la propriété [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) pour obtenir le type de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="df5fc-166">After an add-in is activated, you can use the [mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) property to obtain the currently selected item in Outlook, and the [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property to obtain the type of the current item.</span></span>

<span data-ttu-id="df5fc-167">Vous pouvez éventuellement utiliser l’attribut pour spécifier la classe de message de l’élément et l’attribut pour spécifier si la règle doit être true lorsque l’élément est une sous-classe de la classe `ItemClass` `IncludeSubClasses` spécifiée. </span><span class="sxs-lookup"><span data-stu-id="df5fc-167">You can optionally use the `ItemClass` attribute to specify the message class of the item, and the `IncludeSubClasses` attribute to specify whether the rule should be **true** when the item is a subclass of the specified class.</span></span>

<span data-ttu-id="df5fc-168">Pour plus d’informations sur les classes de message, reportez-vous à la rubrique [Types d’éléments et classes de messages](/office/vba/outlook/Concepts/Forms/item-types-and-message-classes).</span><span class="sxs-lookup"><span data-stu-id="df5fc-168">For more information about message classes, see [Item Types and Message Classes](/office/vba/outlook/Concepts/Forms/item-types-and-message-classes).</span></span>

<span data-ttu-id="df5fc-169">L’exemple suivant est une règle **ItemIs** qui permet aux utilisateurs de voir le add-in dans la barre de Outlook lorsque l’utilisateur lit un message.</span><span class="sxs-lookup"><span data-stu-id="df5fc-169">The following example is an **ItemIs** rule that lets users see the add-in in the Outlook add-in bar when the user is reading a message.</span></span>

```xml
<Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
```

<span data-ttu-id="df5fc-170">L’exemple suivant illustre une règle **ItemIs** permettant aux utilisateurs d’afficher le complément dans la barre de compléments Outlook lorsqu’ils lisent un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="df5fc-170">The following example is an **ItemIs** rule that lets users see the add-in in the Outlook add-in bar when the user is reading a message or appointment.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
</Rule>
```


## <a name="itemhasattachment-rule"></a><span data-ttu-id="df5fc-171">Règle ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="df5fc-171">ItemHasAttachment rule</span></span>


<span data-ttu-id="df5fc-172">Le `ItemHasAttachment` type complexe définit une règle qui vérifie si l’élément sélectionné contient une pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="df5fc-172">The `ItemHasAttachment` complex type defines a rule that checks if the selected item contains an attachment.</span></span>

```xml
<Rule xsi:type="ItemHasAttachment" />
```


## <a name="itemhasknownentity-rule"></a><span data-ttu-id="df5fc-173">Règle ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="df5fc-173">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="df5fc-174">Avant qu’un élément ne soit mis à la disposition d’un complément, le serveur l’examine pour déterminer si l’objet ou le corps contient du texte susceptible de correspondre à l’une des entités connues.</span><span class="sxs-lookup"><span data-stu-id="df5fc-174">Before an item is made available to an add-in, the server examines it to determine whether the subject and body contain any text that is likely to be one of the known entities.</span></span> <span data-ttu-id="df5fc-175">Si l’une de ces entités est trouvée, elle est placée dans une collection d’entités connues accessibles à l’aide de la ou de la `getEntities` `getEntitiesByType` méthode de cet élément.</span><span class="sxs-lookup"><span data-stu-id="df5fc-175">If any of these entities are found, it is placed in a collection of known entities that you access by using the `getEntities` or `getEntitiesByType` method of that item.</span></span>

<span data-ttu-id="df5fc-176">Vous pouvez spécifier une règle à l’aide de celle qui affiche votre add-in lorsqu’une entité du type spécifié `ItemHasKnownEntity` est présente dans l’élément.</span><span class="sxs-lookup"><span data-stu-id="df5fc-176">You can specify a rule by using `ItemHasKnownEntity` that shows your add-in when an entity of the specified type is present in the item.</span></span> <span data-ttu-id="df5fc-177">Vous pouvez spécifier les entités connues suivantes dans `EntityType` l’attribut d’une `ItemHasKnownEntity` règle.</span><span class="sxs-lookup"><span data-stu-id="df5fc-177">You can specify the following known entities in the `EntityType` attribute of an `ItemHasKnownEntity` rule.</span></span>

- <span data-ttu-id="df5fc-178">Address</span><span class="sxs-lookup"><span data-stu-id="df5fc-178">Address</span></span>
- <span data-ttu-id="df5fc-179">Contact</span><span class="sxs-lookup"><span data-stu-id="df5fc-179">Contact</span></span>
- <span data-ttu-id="df5fc-180">EmailAddress</span><span class="sxs-lookup"><span data-stu-id="df5fc-180">EmailAddress</span></span>
- <span data-ttu-id="df5fc-181">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="df5fc-181">MeetingSuggestion</span></span>
- <span data-ttu-id="df5fc-182">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="df5fc-182">PhoneNumber</span></span>
- <span data-ttu-id="df5fc-183">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="df5fc-183">TaskSuggestion</span></span>
- <span data-ttu-id="df5fc-184">URL</span><span class="sxs-lookup"><span data-stu-id="df5fc-184">URL</span></span>

<span data-ttu-id="df5fc-185">Vous pouvez éventuellement inclure une expression régulière dans l’attribut afin que votre add-in ne s’affiche qu’en présence d’une entité qui correspond à `RegularExpression` l’expression régulière.</span><span class="sxs-lookup"><span data-stu-id="df5fc-185">You can optionally include a regular expression in the `RegularExpression` attribute so that your add-in is only shown when an entity that matches the regular expression in present.</span></span> <span data-ttu-id="df5fc-186">Pour obtenir des correspondances avec des expressions régulières spécifiées dans les règles, vous pouvez utiliser la ou la méthode pour `ItemHasKnownEntity` `getRegExMatches` l’élément Outlook actuellement `getFilteredEntitiesByName` sélectionné.</span><span class="sxs-lookup"><span data-stu-id="df5fc-186">To obtain matches to regular expressions specified in `ItemHasKnownEntity` rules, you can use the `getRegExMatches` or `getFilteredEntitiesByName` method for the currently selected Outlook item.</span></span>

<span data-ttu-id="df5fc-187">L’exemple suivant montre une collection d’éléments qui montrent le add-in lorsque l’une des entités connues spécifiées est présente `Rule` dans le message.</span><span class="sxs-lookup"><span data-stu-id="df5fc-187">The following example shows a collection of `Rule` elements that show the add-in when one of the specified well-known entities is present in the message.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="TaskSuggestion" />
</Rule>
```

<span data-ttu-id="df5fc-188">L’exemple suivant montre une règle avec un attribut qui active le module lorsqu’une URL contenant le `ItemHasKnownEntity` mot « contoso » est présente `RegularExpression` dans un message.</span><span class="sxs-lookup"><span data-stu-id="df5fc-188">The following example shows an `ItemHasKnownEntity` rule with a `RegularExpression` attribute that activates the add-in when a URL that contains the word "contoso" is present in a message.</span></span>


```xml
<Rule xsi:type="ItemHasKnownEntity" EntityType="Url" RegularExpression="contoso" />
```

<span data-ttu-id="df5fc-189">Pour plus d’informations sur les entités dans les règles d’activation, voir [Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues](match-strings-in-an-item-as-well-known-entities.md).</span><span class="sxs-lookup"><span data-stu-id="df5fc-189">For more information about entities in activation rules, see [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md).</span></span>


## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="df5fc-190">Règle ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="df5fc-190">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="df5fc-191">Le type complexe définit une règle qui utilise une expression régulière pour correspondre au contenu de la propriété spécifiée `ItemHasRegularExpressionMatch` d’un élément.</span><span class="sxs-lookup"><span data-stu-id="df5fc-191">The `ItemHasRegularExpressionMatch` complex type defines a rule that uses a regular expression to match the contents of the specified property of an item.</span></span> <span data-ttu-id="df5fc-192">Si le texte correspondant à l’expression régulière est trouvé dans la propriété spécifiée de l’élément, Outlook active la barre de compléments et affiche le complément.</span><span class="sxs-lookup"><span data-stu-id="df5fc-192">If text that matches the regular expression is found in the specified property of the item, Outlook activates the add-in bar and displays the add-in.</span></span> <span data-ttu-id="df5fc-193">Vous pouvez utiliser la ou la méthode de l’objet qui représente l’élément actuellement sélectionné pour obtenir des `getRegExMatches` `getRegExMatchesByName` correspondances pour l’expression régulière spécifiée.</span><span class="sxs-lookup"><span data-stu-id="df5fc-193">You can use the `getRegExMatches` or `getRegExMatchesByName` method of the object that represents the currently selected item to obtain matches for the specified regular expression.</span></span>

<span data-ttu-id="df5fc-194">L’exemple suivant montre un exemple qui active le add-in lorsque le corps de l’élément sélectionné contient « apple », « premier » ou « coco » en ignorant la `ItemHasRegularExpressionMatch` cas.</span><span class="sxs-lookup"><span data-stu-id="df5fc-194">The following example shows an `ItemHasRegularExpressionMatch` that activates the add-in when the body of the selected item contains "apple", "banana", or "coconut", ignoring case.</span></span>

```xml
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
```

<span data-ttu-id="df5fc-195">Pour plus d’informations sur l’utilisation de la règle, voir Utiliser des règles d’activation d’expression régulière pour afficher `ItemHasRegularExpressionMatch` [Outlook complément.](use-regular-expressions-to-show-an-outlook-add-in.md)</span><span class="sxs-lookup"><span data-stu-id="df5fc-195">For more information about using the `ItemHasRegularExpressionMatch` rule, see [Use regular expression activation rules to show an Outlook add-in](use-regular-expressions-to-show-an-outlook-add-in.md).</span></span>


## <a name="rulecollection-rule"></a><span data-ttu-id="df5fc-196">Règle RuleCollection</span><span class="sxs-lookup"><span data-stu-id="df5fc-196">RuleCollection rule</span></span>


<span data-ttu-id="df5fc-197">Le `RuleCollection` type complexe combine plusieurs règles en une seule règle.</span><span class="sxs-lookup"><span data-stu-id="df5fc-197">The `RuleCollection` complex type combines multiple rules into a single rule.</span></span> <span data-ttu-id="df5fc-198">Vous pouvez spécifier si les règles de la collection doivent être combinées avec un OU logique ou un AND logique à l’aide de `Mode` l’attribut.</span><span class="sxs-lookup"><span data-stu-id="df5fc-198">You can specify whether the rules in the collection should be combined with a logical OR or a logical AND by using the `Mode` attribute.</span></span>

<span data-ttu-id="df5fc-p118">Quand un ET logique est spécifié, un élément doit correspondre à toutes les règles spécifiées dans le regroupement pour permettre l’affichage du complément. Quand un OU logique est spécifié, tout élément qui correspond à l’une des règles spécifiées dans le regroupement permet l’affichage du complément.</span><span class="sxs-lookup"><span data-stu-id="df5fc-p118">When a logical AND is specified, an item must match all the specified rules in the collection to show the add-in. When a logical OR is specified, an item that matches any of the specified rules in the collection will show the add-in.</span></span>

<span data-ttu-id="df5fc-201">Vous pouvez combiner `RuleCollection` des règles pour former des règles complexes.</span><span class="sxs-lookup"><span data-stu-id="df5fc-201">You can combine `RuleCollection` rules to form complex rules.</span></span> <span data-ttu-id="df5fc-202">L’exemple suivant illustre l’activation du complément lorsque l’utilisateur affiche un élément de rendez-vous ou de message et que l’objet ou le corps de l’élément contient une adresse.</span><span class="sxs-lookup"><span data-stu-id="df5fc-202">The following example activates the add-in when the user is viewing an appointment or message item and the subject or body of the item contains an address.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read"/>
  </Rule>
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

<span data-ttu-id="df5fc-203">L’exemple suivant illustre l’activation du complément lorsque l’utilisateur compose un message ou affiche un rendez-vous, et que l’objet ou le corps du rendez-vous contient une adresse.</span><span class="sxs-lookup"><span data-stu-id="df5fc-203">The following example activates the add-in when the user is composing a message, or when the user is viewing an appointment and the subject or body of the appointment contains an address.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="Or"> 
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" /> 
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
  </Rule> 
</Rule>
```


## <a name="limits-for-rules-and-regular-expressions"></a><span data-ttu-id="df5fc-204">Limites pour les règles et les expressions régulières</span><span class="sxs-lookup"><span data-stu-id="df5fc-204">Limits for rules and regular expressions</span></span>


<span data-ttu-id="df5fc-205">Pour fournir une expérience satisfaisante avec les compléments Outlook, vous devez vous conformer aux directives d’activation et d’utilisation des API.</span><span class="sxs-lookup"><span data-stu-id="df5fc-205">To provide a satisfactory experience with Outlook add-ins, you should adhere to the activation and API usage guidelines.</span></span> <span data-ttu-id="df5fc-206">Le tableau suivant présente des limites générales pour les expressions régulières et les règles, mais il existe des règles spécifiques pour différentes applications.</span><span class="sxs-lookup"><span data-stu-id="df5fc-206">The following table shows general limits for regular expressions and rules but there are specific rules for different applications.</span></span> <span data-ttu-id="df5fc-207">Pour plus d’informations, voir [Limites d’activation et d’API JavaScript des compléments Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) et [Résoudre les problèmes d’activation des compléments Outlook](troubleshoot-outlook-add-in-activation.md).</span><span class="sxs-lookup"><span data-stu-id="df5fc-207">For more information, see [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) and [Troubleshoot Outlook add-in activation](troubleshoot-outlook-add-in-activation.md).</span></span>

<br/>

|<span data-ttu-id="df5fc-208">**Élément de complément**</span><span class="sxs-lookup"><span data-stu-id="df5fc-208">**Add-in element**</span></span>|<span data-ttu-id="df5fc-209">**Conseils**</span><span class="sxs-lookup"><span data-stu-id="df5fc-209">**Guidelines**</span></span>|
|:-----|:-----|
|<span data-ttu-id="df5fc-210">Taille de manifeste</span><span class="sxs-lookup"><span data-stu-id="df5fc-210">Manifest Size</span></span>|<span data-ttu-id="df5fc-211">Inférieur à 256 Ko.</span><span class="sxs-lookup"><span data-stu-id="df5fc-211">No larger than 256 KB.</span></span>|
|<span data-ttu-id="df5fc-212">Règles</span><span class="sxs-lookup"><span data-stu-id="df5fc-212">Rules</span></span>|<span data-ttu-id="df5fc-213">Pas plus de 15 règles.</span><span class="sxs-lookup"><span data-stu-id="df5fc-213">No more than 15 rules.</span></span>|
|<span data-ttu-id="df5fc-214">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="df5fc-214">ItemHasKnownEntity</span></span>|<span data-ttu-id="df5fc-215">Un riche client Outlook appliquera la règle au premier mégaoctet du corps, mais pas au reste.</span><span class="sxs-lookup"><span data-stu-id="df5fc-215">An Outlook rich client will apply the rule against the first 1 MB of the body, and not to the rest of the body.</span></span>|
|<span data-ttu-id="df5fc-216">Expressions régulières</span><span class="sxs-lookup"><span data-stu-id="df5fc-216">Regular Expressions</span></span>|<span data-ttu-id="df5fc-217">Pour les règles ItemHasKnownEntity ou ItemHasRegularExpressionMatch pour toutes Outlook applications :</span><span class="sxs-lookup"><span data-stu-id="df5fc-217">For ItemHasKnownEntity or ItemHasRegularExpressionMatch rules for all Outlook applications:</span></span><br><ul><li><span data-ttu-id="df5fc-p121">Ne spécifiez pas plus de 5 expressions régulières dans les règles d’activation pour un complément Outlook. Vous ne pouvez pas installer de complément si vous dépassez cette limite.</span><span class="sxs-lookup"><span data-stu-id="df5fc-p121">Specify no more than 5 regular expressions in activation rules for an Outlook add-in. You cannot install an add-in if you exceed that limit.</span></span></li><li><span data-ttu-id="df5fc-220">Spécifiez des expressions régulières dont les résultats sont renvoyés par l’appel de la méthode <b>getRegExMatches</b> dans les 50 premières correspondances.</span><span class="sxs-lookup"><span data-stu-id="df5fc-220">Specify regular expressions whose anticipated results are returned by the <b>getRegExMatches</b> method call within the first 50 matches.</span></span> </li><li><span data-ttu-id="df5fc-221">Spécifiez des assertions avant dans les expressions régulières, mais pas d’assertions arrière, `(?<=text)`, ni d’assertions arrière négatives `(?<!text)`.</span><span class="sxs-lookup"><span data-stu-id="df5fc-221">Specify look-ahead assertions in regular expressions, but not look-behind, `(?<=text)`, and negative look-behind `(?<!text)`.</span></span></li><li><span data-ttu-id="df5fc-222">Spécifiez des expressions régulières dont la correspondance ne dépasse pas les limites figurant dans le tableau ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="df5fc-222">Specify regular expressions whose match does not exceed the limits in the table below.</span></span><br/><br/><table><tr><th><span data-ttu-id="df5fc-223">Limite de longueur d’une correspondance d’expression régulière</span><span class="sxs-lookup"><span data-stu-id="df5fc-223">Limit on length of a regex match</span></span></th><th><span data-ttu-id="df5fc-224">Clients riches Outlook</span><span class="sxs-lookup"><span data-stu-id="df5fc-224">Outlook rich clients</span></span></th><th><span data-ttu-id="df5fc-225">Outlook sur iOS et Android</span><span class="sxs-lookup"><span data-stu-id="df5fc-225">Outlook on iOS and Android</span></span></th></tr><tr><td><span data-ttu-id="df5fc-226">Corps d’élément en texte brut</span><span class="sxs-lookup"><span data-stu-id="df5fc-226">Item body is plain text</span></span></td><td><span data-ttu-id="df5fc-227">1,5 Ko</span><span class="sxs-lookup"><span data-stu-id="df5fc-227">1.5 KB</span></span></td><td><span data-ttu-id="df5fc-228">3 Ko</span><span class="sxs-lookup"><span data-stu-id="df5fc-228">3 KB</span></span></td></tr><tr><td><span data-ttu-id="df5fc-229">Corps d’élément en HTML</span><span class="sxs-lookup"><span data-stu-id="df5fc-229">Item body it HTML</span></span></td><td><span data-ttu-id="df5fc-230">3 Ko</span><span class="sxs-lookup"><span data-stu-id="df5fc-230">3 KB</span></span></td><td><span data-ttu-id="df5fc-231">3 Ko</span><span class="sxs-lookup"><span data-stu-id="df5fc-231">3 KB</span></span></td></tr></table>|

## <a name="see-also"></a><span data-ttu-id="df5fc-232">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="df5fc-232">See also</span></span>

- [<span data-ttu-id="df5fc-233">Créer des compléments Outlook pour les formulaires de composition</span><span class="sxs-lookup"><span data-stu-id="df5fc-233">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)
- [<span data-ttu-id="df5fc-234">Limites pour l’activation et l’API JavaScript pour les compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="df5fc-234">Limits for activation and JavaScript API for Outlook add-ins</span></span>](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [<span data-ttu-id="df5fc-235">Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook</span><span class="sxs-lookup"><span data-stu-id="df5fc-235">Use regular expression activation rules to show an Outlook add-in</span></span>](use-regular-expressions-to-show-an-outlook-add-in.md)
- [<span data-ttu-id="df5fc-236">Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues</span><span class="sxs-lookup"><span data-stu-id="df5fc-236">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)
    
