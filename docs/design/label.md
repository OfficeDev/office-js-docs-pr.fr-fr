---
title: Composant d’étiquette dans la structure de l’interface utilisateur d’Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: e9d6e9eaca918068b682725ee9236f6539641fa0
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437240"
---
# <a name="label-component-in-office-ui-fabric"></a><span data-ttu-id="bf04b-102">Composant d’étiquette dans la structure de l’interface utilisateur d’Office</span><span class="sxs-lookup"><span data-stu-id="bf04b-102">Label component in Office UI Fabric</span></span>

<span data-ttu-id="bf04b-p101">Utilisez des étiquettes pour nommer ou donner un titre à un composant ou un groupe de composants. Associées à un autre composant ou groupe de composants, les étiquettes doivent se trouver à proximité des composants ou des groupes associés. Certains composants ont des étiquettes prédéfinies comme les listes déroulantes ou les boutons bascule.</span><span class="sxs-lookup"><span data-stu-id="bf04b-p101">Use labels to name or title a component or group of components. When paired with another component or group of components, labels should be in close proximity to the related components or groups. Some components have predefined labels, such as a drop-down or toggle.</span></span>
  
#### <a name="example-label-in-a-task-pane"></a><span data-ttu-id="bf04b-106">Exemple : Étiquette dans un volet de tâches</span><span class="sxs-lookup"><span data-stu-id="bf04b-106">Example: Label in a task pane</span></span>

![Image illustrant l’étiquette](../images/overview-with-app-label.png)

## <a name="best-practices"></a><span data-ttu-id="bf04b-108">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="bf04b-108">Best practices</span></span>

|<span data-ttu-id="bf04b-109">**À faire**</span><span class="sxs-lookup"><span data-stu-id="bf04b-109">**Do**</span></span>|<span data-ttu-id="bf04b-110">**À ne pas faire**</span><span class="sxs-lookup"><span data-stu-id="bf04b-110">**Don't**</span></span>|
|:------------|:--------------|
|<span data-ttu-id="bf04b-111">Utilisez la casse pour une phrase, par exemple **Prénom**.</span><span class="sxs-lookup"><span data-stu-id="bf04b-111">Use sentence casing, for example **First name**.</span></span>|<span data-ttu-id="bf04b-112">N’utilisez pas la casse pour un titre, par exemple **Prénom**.</span><span class="sxs-lookup"><span data-stu-id="bf04b-112">Don’t use title casing, for example **First Name**.</span></span>|
|<span data-ttu-id="bf04b-113">Soyez court et concis.</span><span class="sxs-lookup"><span data-stu-id="bf04b-113">Be short and concise.</span></span>|<span data-ttu-id="bf04b-114">N’utilisez pas de phrases complètes ni de signes de ponctuation complexes comme les virgules ou les points-virgules.</span><span class="sxs-lookup"><span data-stu-id="bf04b-114">Don’t use full sentences or complex punctuation, such as colons or semicolons.</span></span>|
|<span data-ttu-id="bf04b-115">Lorsque vous ajoutez une étiquette à des composants, utilisez un nom ou une locution nominale courte comme texte d’étiquette.</span><span class="sxs-lookup"><span data-stu-id="bf04b-115">When adding a label to components, use a noun or short noun phrase as the label text.</span></span>| |

## <a name="variants"></a><span data-ttu-id="bf04b-116">Variantes</span><span class="sxs-lookup"><span data-stu-id="bf04b-116">Variants</span></span>

|<span data-ttu-id="bf04b-117">**Variation**</span><span class="sxs-lookup"><span data-stu-id="bf04b-117">**Variation**</span></span>|<span data-ttu-id="bf04b-118">**Description**</span><span class="sxs-lookup"><span data-stu-id="bf04b-118">**Description**</span></span>|<span data-ttu-id="bf04b-119">**Exemple**</span><span class="sxs-lookup"><span data-stu-id="bf04b-119">**Example**</span></span>|
|:------------|:--------------|:----------|
|<span data-ttu-id="bf04b-120">**Étiquette par défaut**</span><span class="sxs-lookup"><span data-stu-id="bf04b-120">**Default label**</span></span>|<span data-ttu-id="bf04b-121">À utiliser pour les étiquettes standard.</span><span class="sxs-lookup"><span data-stu-id="bf04b-121">Use for standard labels.</span></span>|![Image de l’étiquette par défaut](../images/label.png)<br/>|
|<span data-ttu-id="bf04b-123">**Étiquette désactivée**</span><span class="sxs-lookup"><span data-stu-id="bf04b-123">**Disabled label**</span></span>|<span data-ttu-id="bf04b-124">À utiliser lorsque le composant associé est désactivé.</span><span class="sxs-lookup"><span data-stu-id="bf04b-124">Use when the related component is disabled.</span></span>|![Image d’étiquette désactivée](../images/label-disabled.png)<br/>|
|<span data-ttu-id="bf04b-126">**Étiquette requise**</span><span class="sxs-lookup"><span data-stu-id="bf04b-126">**Required label**</span></span>|<span data-ttu-id="bf04b-127">À utiliser lorsque le composant associé est requis.</span><span class="sxs-lookup"><span data-stu-id="bf04b-127">Use when the related component is required.</span></span>|![Image d’étiquette requise](../images/label-required.png)<br/>|

## <a name="implementation"></a><span data-ttu-id="bf04b-129">Implémentation</span><span class="sxs-lookup"><span data-stu-id="bf04b-129">Implementation</span></span>

<span data-ttu-id="bf04b-130">Pour plus d’informations, reportez-vous à [Étiquette](https://dev.office.com/fabric#/components/label) et [Démarrer avec un exemple de code React de la structure](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).</span><span class="sxs-lookup"><span data-stu-id="bf04b-130">For details, see [Label](https://dev.office.com/fabric#/components/label) and [Getting started with Fabric React code sample](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).</span></span>

## <a name="see-also"></a><span data-ttu-id="bf04b-131">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="bf04b-131">See also</span></span>

- [<span data-ttu-id="bf04b-132">Modèles de conception UX</span><span class="sxs-lookup"><span data-stu-id="bf04b-132">UX Design Patterns</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [<span data-ttu-id="bf04b-133">Office UI Fabric dans des compléments Office</span><span class="sxs-lookup"><span data-stu-id="bf04b-133">Office UI Fabric in Office Add-ins</span></span>](office-ui-fabric.md)
