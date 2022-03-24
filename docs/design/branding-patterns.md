---
title: Instructions de conception de modèles de personnalisation pour les compléments Office
description: Découvrez comment  brander votre Office tout en restant compatible avec la conception visuelle de Office.
ms.date: 06/26/2018
ms.localizationpriority: medium
ms.openlocfilehash: 529daa6b6e78f1cff1ea7ea3c61efb40f8e4eb60
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63742941"
---
# <a name="branding-patterns"></a>Modèles de personnalisation

Ces modèles offrent une visibilité de marque et un contexte aux utilisateurs de votre add-in.

## <a name="best-practices"></a>Meilleures pratiques

|À faire |À ne pas faire|
|:---- |:----|
| Utilisez des composants d’interface utilisateur familiers en même temps que des caractéristiques de votre marque, comme par exemple une typographie et des couleurs typiques. | N’inventez pas des nouveaux composants d’interface utilisateur qui s’opposent aux éléments d’interface utilisateur établis pour Office. |
| Placez la personnalisation de marque pour le complément dans une barre de marque en pied de page en bas de votre interface utilisateur. | Ne répétez pas le nom du volet Office dans une barre de marque immédiatement adjacentes dans la partie supérieure de votre interface utilisateur. |
| Utilisez les éléments de marque avec parcimonie. Intégrez votre solution à Office pour qu’elle soit complémentaire. | N’insérez pas trop d’éléments de personnalisation dans l’interface utilisateur Office, cela risque de détourner l’attention des clients et de les rendre confus. |
| Assurez que votre solution soit facilement reconnaissable et assurez la continuité de vos écrans avec des éléments visuels cohérentes. | Ne masquez pas votre solution avec des éléments visuels inconnus et appliqués de manière incohérente. |
| Créez une connexion avec un service ou une entreprise parent pour vous assurer que les clients connaissent et apprécient votre solution. | Ne forcez pas les clients à apprendre un nouveau concept de marque s’il existe déjà une relation utile et compréhensible qui peut être utilisée pour créer la confiance et ajouter de la valeur. |

Appliquer les modèles et les composants suivants le cas échéant pour permettre aux utilisateurs de comprendre et utiliser toute l’utilité de votre complément.

## <a name="brand-bar"></a>Barre de marque

La barre de marque est un espace dans le pied de page où vous pouvez inclure le nom de la marque et le logo. Elle sert également de lien vers le site Web de votre marque et d’emplacement d’accès facultatif.

![Barre de marque affichée dans le volet Des tâches d’un Office application de bureau.](../images/add-in-brand-bar.png)

## <a name="splash-screen"></a>Écran de démarrage

Utilisez cet écran pour afficher votre personnalisation pendant que le complément est en cours de chargement ou lors de la transition entre les différents états de l’interface utilisateur.

![Écran de marque affiché dans le volet Des tâches d’un Office application de bureau.](../images/add-in-splash-screen.png)
