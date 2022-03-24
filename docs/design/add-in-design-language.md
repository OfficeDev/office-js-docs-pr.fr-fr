---
title: Langage de création d’un complément Office
description: Découvrez comment faire en sorte que votre Office soit visuellement compatible avec Office.
ms.date: 05/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: a7623b305ec7857957ce1215574ea5cdee70b3d7
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63742794"
---
# <a name="office-add-in-design-language"></a>Langage de création d’un complément Office

Le langage de création d’Office est un système visuel clair et simple qui garantit la cohérence entre expériences. Il contient un ensemble d’éléments visuels qui définissent les interfaces Office, y compris :

- Police standard
- Palette de couleurs courantes
- Ensemble de tailles typographiques et pondérations
- Instructions relatives aux icônes
- Éléments d’icône partagée
- Définitions d’animation
- Composants courants

[Fluent’interface utilisateur](../design/add-in-design.md) est l’infrastructure frontale officielle pour la création avec Office langage de conception. L Fluent’interface utilisateur est facultative, mais il s’agit du moyen le plus rapide pour vous assurer que vos Office. Tirez parti de Fluent’interface utilisateur pour concevoir et créer des compléments qui complètent Office.

De nombreux compléments d’Office sont associés à une marque préexistante. Vous pouvez conserver une marque forte et son langage de composant ou visuel dans votre complément. Recherchez les opportunités pour conserver votre propre langage visuel lors de l’intégration avec Office. Pensez à des moyens de remplacer les couleurs Office, la typographie, les icônes ou d’autres éléments stylistiques par des éléments de votre marque. Pensez à des moyens de suivre des dispositions de complément ou des modèles de conception de l’expérience utilisateur courants tout en insérant des contrôles et des composants que vos clients connaissent.

L’insertion d’une interface utilisateur HTML de marque importante à l’intérieur d’Office peut créer des dissonances pour les clients. Trouvez un équilibre qui s’adapte en toute transparence dans Office mais qui s’aligne aussi clairement sur votre marque parent ou de service. Lorsqu’un complément ne s’adapte pas à Office, c’est souvent en raison d’une incompatibilité des éléments stylistiques. Par exemple, la typographie est trop grande et en dehors de la grille, les couleurs sont particulièrement criardes ou contrastées, ou les animations sont superflues et se comportent différemment par rapport à Office. L’apparence et le comportement des contrôles ou des composants dévient trop des normes d’Office.
