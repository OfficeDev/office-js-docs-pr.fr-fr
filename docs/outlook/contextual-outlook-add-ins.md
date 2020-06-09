---
title: Compléments Outlook contextuels
description: Lancer des tâches liées à un message sans laisser le message lui-même pour faciliter et enrichir l'expérience utilisateur.
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: c2cfbc1019048bb02186521c2cb81ed832934a8d
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608956"
---
# <a name="contextual-outlook-add-ins"></a><span data-ttu-id="cb0ba-103">Compléments Outlook contextuels</span><span class="sxs-lookup"><span data-stu-id="cb0ba-103">Contextual Outlook add-ins</span></span>

<span data-ttu-id="cb0ba-p101">Les compléments contextuels sont des compléments Outlook qui s’activent en fonction du texte d’un message ou d’un rendez-vous. Grâce aux compléments contextuels, vous pouvez initier des tâches associées à un message sans avoir à quitter ce dernier. L’expérience utilisateur en est ainsi facilitée et enrichie.</span><span class="sxs-lookup"><span data-stu-id="cb0ba-p101">Contextual add-ins are Outlook add-ins that activate based on text in a message or appointment. By using contextual add-ins, a user can initiate tasks related to a message without leaving the message itself, which results in an easier and richer user experience.</span></span>

<span data-ttu-id="cb0ba-106">Voici quelques exemples de compléments contextuels :</span><span class="sxs-lookup"><span data-stu-id="cb0ba-106">The following are examples of contextual add-ins:</span></span>

- <span data-ttu-id="cb0ba-107">Choix d’une adresse à ouvrir dans un plan du lieu.</span><span class="sxs-lookup"><span data-stu-id="cb0ba-107">Choosing an address to open a map of the location.</span></span>
- <span data-ttu-id="cb0ba-108">Choix d’une chaîne ouvrant un complément de suggestion de réunion.</span><span class="sxs-lookup"><span data-stu-id="cb0ba-108">Choosing a string that opens a meeting suggestion add-in.</span></span>
- <span data-ttu-id="cb0ba-109">Choisir un numéro de téléphone permet de l’ajouter à vos contacts.</span><span class="sxs-lookup"><span data-stu-id="cb0ba-109">Choosing a phone number to add to your contacts.</span></span>


> [!NOTE]
> <span data-ttu-id="cb0ba-110">Les compléments contextuels ne sont pas disponibles actuellement dans Outlook pour Android et iOS.</span><span class="sxs-lookup"><span data-stu-id="cb0ba-110">Contextual add-ins are not currently available in Outlook on Android and iOS.</span></span> <span data-ttu-id="cb0ba-111">Cette fonctionnalité sera disponible ultérieurement.</span><span class="sxs-lookup"><span data-stu-id="cb0ba-111">This functionality will be made available in the future.</span></span>
>
> <span data-ttu-id="cb0ba-112">La prise en charge de cette fonctionnalité a été introduite dans l’ensemble de conditions requises 1.6.</span><span class="sxs-lookup"><span data-stu-id="cb0ba-112">Support for this feature was introduced in requirement set 1.6.</span></span> <span data-ttu-id="cb0ba-113">Voir [les clients et les plateformes](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.</span><span class="sxs-lookup"><span data-stu-id="cb0ba-113">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="how-to-make-a-contextual-add-in"></a><span data-ttu-id="cb0ba-114">Création d’un complément contextuel</span><span class="sxs-lookup"><span data-stu-id="cb0ba-114">How to make a contextual add-in</span></span>

<span data-ttu-id="cb0ba-115">Le manifeste d’un complément contextuel doit inclure un élément [ExtensionPoint](../reference/manifest/extensionpoint.md#detectedentity) avec une attribut `xsi:type` défini sur `DetectedEntity`.</span><span class="sxs-lookup"><span data-stu-id="cb0ba-115">A contextual add-in's manifest must include an [ExtensionPoint](../reference/manifest/extensionpoint.md#detectedentity) element with an `xsi:type` attribute set to `DetectedEntity`.</span></span> <span data-ttu-id="cb0ba-116">Au sein de l’élément **ExtensionPoint**, le complément spécifie les entités ou l’expression régulière qui peuvent l’activer.</span><span class="sxs-lookup"><span data-stu-id="cb0ba-116">Within the **ExtensionPoint** element, the add-in specifies the entities or regular expression that can activate it.</span></span> <span data-ttu-id="cb0ba-117">Si une entité est spécifiée, il peut s’agir d’une des propriétés de l’objet [Entités](/javascript/api/outlook/office.entities).</span><span class="sxs-lookup"><span data-stu-id="cb0ba-117">If an entity is specified, the entity can be any of the properties in the [Entities](/javascript/api/outlook/office.entities) object.</span></span>

<span data-ttu-id="cb0ba-118">Par conséquent, le manifeste du complément doit contenir un type de règle **ItemHasKnownEntity** ou **Itemhasregularexpressionmatch**.</span><span class="sxs-lookup"><span data-stu-id="cb0ba-118">Thus, the add-in manifest must contain a rule of type **ItemHasKnownEntity** or **ItemHasRegularExpressionMatch**.</span></span> <span data-ttu-id="cb0ba-119">L’exemple suivant montre comment spécifier qu’un complément doit s’activer sur les messages comportant une entité détectée telle qu’un numéro de téléphone :</span><span class="sxs-lookup"><span data-stu-id="cb0ba-119">The following example shows how to specify that an add-in should activate on messages with a detected entity that is a phone number:</span></span>

```XML
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="contextLabel" />
  <!--If you opt to include RequestedHeight, it must be between 140px to 450px, inclusive.-->
  <!--<RequestedHeight>360</RequestedHeight>-->
  <SourceLocation resid="detectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="PhoneNumber" Highlight="all" />
  </Rule>
</ExtensionPoint>
```

<span data-ttu-id="cb0ba-120">Une fois qu’un complément contextuel est associé à un compte, il démarre automatiquement lorsque l’utilisateur clique sur une expression régulière ou une entité mise en surbrillance.</span><span class="sxs-lookup"><span data-stu-id="cb0ba-120">After a contextual add-in is associated with an account, it will automatically start when the user clicks a highlighted entity or regular expression.</span></span> <span data-ttu-id="cb0ba-121">Pour plus d’informations sur les expressions régulières pour les compléments Outlook, reportez-vous à l’article [Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook](use-regular-expressions-to-show-an-outlook-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="cb0ba-121">For more information about regular expressions for Outlook add-ins, see [Use regular expression activation rules to show an Outlook add-in](use-regular-expressions-to-show-an-outlook-add-in.md).</span></span>

<span data-ttu-id="cb0ba-122">Il existe plusieurs restrictions sur les compléments contextuels :</span><span class="sxs-lookup"><span data-stu-id="cb0ba-122">There are several restrictions on contextual add-ins:</span></span>

- <span data-ttu-id="cb0ba-123">Un complément contextuel ne peut exister que dans des compléments de lecture (pas dans des compléments de composition).</span><span class="sxs-lookup"><span data-stu-id="cb0ba-123">A contextual add-in can only exist in read add-ins (not compose add-ins).</span></span>
- <span data-ttu-id="cb0ba-124">Vous ne pouvez pas spécifier la couleur de l’entité en surbrillance.</span><span class="sxs-lookup"><span data-stu-id="cb0ba-124">You cannot specify the color of the highlighted entity.</span></span>
- <span data-ttu-id="cb0ba-125">Si une entité n’est pas en surbrillance, elle ne lancera pas de complément contextuel dans une carte.</span><span class="sxs-lookup"><span data-stu-id="cb0ba-125">An entity that is not highlighted will not launch a contextual add-in in a card.</span></span>

<span data-ttu-id="cb0ba-126">Une entité ou une expression régulière qui n’est pas mise en surbrillance ne permettant pas le lancement d’un complément contextuel, les compléments doivent inclure au moins un élément `Rule` avec l’attribut `Highlight` défini sur `all`.</span><span class="sxs-lookup"><span data-stu-id="cb0ba-126">Because an entity or regular expression that is not highlighted will not launch a contextual add-in, add-ins must include at least one `Rule` element with the `Highlight` attribute set to `all`.</span></span>

> [!NOTE]
> <span data-ttu-id="cb0ba-p107">Les types d’entité `EmailAddress` et `Url` ne prennent pas en charge la mise en surbrillance. Ils ne peuvent donc pas être utilisés pour lancer un complément contextuel. Ils peuvent toutefois être combinés dans un type de règle `RuleCollection` comme un critère d’activation supplémentaire.</span><span class="sxs-lookup"><span data-stu-id="cb0ba-p107">The `EmailAddress` and `Url` entity types do not support highlighting, so they cannot be used to launch a contextual add-in. They can however be combined in a `RuleCollection` rule type as an additional activation criteria.</span></span>

## <a name="how-to-launch-a-contextual-add-in"></a><span data-ttu-id="cb0ba-129">Lancement d’un complément contextuel</span><span class="sxs-lookup"><span data-stu-id="cb0ba-129">How to launch a contextual add-in</span></span>

<span data-ttu-id="cb0ba-p108">Un utilisateur lance un complément contextuel par le biais du texte, une entité connue ou une expression régulière du développeur. En règle générale, un utilisateur identifie un complément contextuel car l’entité est mise en surbrillance. L’exemple suivant montre comment la mise en surbrillance s’affiche dans un message. Ici, l’entité (une adresse) est colorée en bleu et soulignée avec une ligne bleue en pointillés. Un utilisateur lance le complément contextuel en cliquant sur l’entité en surbrillance.</span><span class="sxs-lookup"><span data-stu-id="cb0ba-p108">A user launches a contextual add-in through text, either a known entity or a developer's regular expression. Typically, a user identifies a contextual add-in because the entity is highlighted. The following example shows how highlighting appears in a message. Here the entity (an address) is colored blue and underlined with a dotted blue line. A user launches the contextual add-in by clicking the highlighted entity.</span></span> 

<span data-ttu-id="cb0ba-135">**Exemple de texte avec l’entité (une adresse) en surbrillance**</span><span class="sxs-lookup"><span data-stu-id="cb0ba-135">**Example of text with highlighted entity (an address)**</span></span>

![Illustre l’entité en surbrillance dans un courrier](../images/outlook-detected-entity-highlight.png)
    
<span data-ttu-id="cb0ba-137">Lorsque plusieurs entités ou compléments contextuels sont présents dans un message, l’interaction avec l’utilisateur a lieu comme suit :</span><span class="sxs-lookup"><span data-stu-id="cb0ba-137">When there are multiple entities or contextual add-ins in a message, there are a few user interaction rules:</span></span>

- <span data-ttu-id="cb0ba-138">S’il existe plusieurs entités, l’utilisateur doit cliquer sur une autre entité pour lancer le complément de celle-ci.</span><span class="sxs-lookup"><span data-stu-id="cb0ba-138">If there are multiple entities, the user has to click a different entity to launch the add-in for it.</span></span>
- <span data-ttu-id="cb0ba-139">Si une entité active plusieurs compléments, chaque complément ouvre un nouvel onglet. L’utilisateur bascule entre les onglets pour passer d’un complément à l’autre. Par exemple, un nom et une adresse peuvent déclencher un complément de téléphone et une carte.</span><span class="sxs-lookup"><span data-stu-id="cb0ba-139">If an entity activates multiple add-ins, each add-in opens a new tab. The user switches between tabs to change between add-ins. For example, a name and address might trigger a phone add-in and a map.</span></span>
- <span data-ttu-id="cb0ba-p109">Si une chaîne unique contient plusieurs entités qui activent plusieurs compléments, la chaîne entière est mise en surbrillance et lorsque l’utilisateur clique sur cette chaîne, tous les compléments concernés par la chaîne s’affichent dans des onglets distincts. Par exemple, une chaîne qui décrit une proposition de réunion dans un restaurant peut activer le complément de suggestion de réunion et un complément d’avis sur des restaurants.</span><span class="sxs-lookup"><span data-stu-id="cb0ba-p109">If a single string contains multiple entities that activate multiple add-ins, the entire string is highlighted, and clicking the string shows all add-ins relevant to the string on separate tabs. For example, a string that describes a proposed meeting at a restaurant might activate the Suggested Meeting add-in and a restaurant rating add-in.</span></span>

## <a name="how-a-contextual-add-in-displays"></a><span data-ttu-id="cb0ba-142">Affichage des compléments contextuels</span><span class="sxs-lookup"><span data-stu-id="cb0ba-142">How a contextual add-in displays</span></span>

<span data-ttu-id="cb0ba-p110">Un complément contextuel activé s’affiche sur une carte, qui est une fenêtre séparée près de l’entité. La carte s’affiche normalement en-dessous de l’entité et le plus centrée possible par rapport à l’entité. S’il n’existe pas suffisamment d’espace en-dessous de l’entité, la carte est placée au-dessus. La capture d’écran suivante illustre l’entité en surbrillance et, dessous, un complément activé (Plans Bing) sur une carte.</span><span class="sxs-lookup"><span data-stu-id="cb0ba-p110">An activated contextual add-in appears in a card, which is a separate window near the entity. The card will normally appear below the entity and centered with respect to the entity as much as possible. If there is not enough room below the entity, the card is placed above it. The following screenshot shows the highlighted entity, and below it, an activated add-in (Bing Maps) in a card.</span></span>

<span data-ttu-id="cb0ba-147">**Exemple d’un complément affiché sur une carte**</span><span class="sxs-lookup"><span data-stu-id="cb0ba-147">**Example of an add-in displayed in a card**</span></span>

![Indique une application contextuelle sur une carte](../images/outlook-detected-entity-card.png)

<span data-ttu-id="cb0ba-149">Pour fermer la carte et quitter le complément, il suffit de cliquer n’importe où en dehors de la carte.</span><span class="sxs-lookup"><span data-stu-id="cb0ba-149">To close the card and the add-in, a user clicks anywhere outside of the card.</span></span>

## <a name="current-contextual-add-ins"></a><span data-ttu-id="cb0ba-150">Compléments contextuels actuels</span><span class="sxs-lookup"><span data-stu-id="cb0ba-150">Current contextual add-ins</span></span>

<span data-ttu-id="cb0ba-151">Les compléments contextuels suivants sont installés par défaut pour les utilisateurs qui utilisent des compléments Outlook :</span><span class="sxs-lookup"><span data-stu-id="cb0ba-151">The following contextual add-ins are installed by default for users with Outlook add-ins:</span></span>

- <span data-ttu-id="cb0ba-152">Plans Bing</span><span class="sxs-lookup"><span data-stu-id="cb0ba-152">Bing Maps</span></span> 
- <span data-ttu-id="cb0ba-153">Réunions suggérées</span><span class="sxs-lookup"><span data-stu-id="cb0ba-153">Suggested Meetings</span></span>

## <a name="see-also"></a><span data-ttu-id="cb0ba-154">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="cb0ba-154">See also</span></span>

- <span data-ttu-id="cb0ba-155">[Complément Outlook : numéro de commande Contoso](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) (exemple de complément contextuel qui est activé en fonction d’une correspondance d’expression régulière)</span><span class="sxs-lookup"><span data-stu-id="cb0ba-155">[Outlook add-in: Contoso Order Number](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) (sample contextual add-in that activates based on a regular expression match)</span></span>
- [<span data-ttu-id="cb0ba-156">Créer votre premier complément Outlook</span><span class="sxs-lookup"><span data-stu-id="cb0ba-156">Write your first Outlook add-in</span></span>](../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="cb0ba-157">Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook</span><span class="sxs-lookup"><span data-stu-id="cb0ba-157">Use regular expression activation rules to show an Outlook add-in</span></span>](use-regular-expressions-to-show-an-outlook-add-in.md)
- [<span data-ttu-id="cb0ba-158">Objet Entités</span><span class="sxs-lookup"><span data-stu-id="cb0ba-158">Entities object</span></span>](/javascript/api/outlook/office.entities)
