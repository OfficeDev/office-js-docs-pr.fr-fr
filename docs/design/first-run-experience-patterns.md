---
title: Modèles de première expérience d’utilisation des complément Office
description: Découvrez les meilleures pratiques pour concevoir des expériences de première Office des modules.
ms.date: 07/08/2018
ms.localizationpriority: medium
ms.openlocfilehash: 43127e2a83c07ae659c6672a57486e5488ad268e
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743246"
---
# <a name="first-run-experience-patterns"></a>Modèles de première expérience d’utilisation

Une première expérience d’utilisation (FRE) correspond à l’introduction d’un utilisateur à votre complément. Une FRE existe quand un utilisateur ouvre un complément pour la première fois et lui fournit des informations sur les fonctions, les fonctionnalités et/ou les avantages du complément. Cette expérience vous permet de modeler la première impression qu’un utilisateur va avoir d’un complément. Elle peut grandement influencer la probabilité qu’il y revienne et continue à utiliser votre complément...

## <a name="best-practices"></a>Meilleures pratiques

Suivez ces meilleures pratiques lors de la création de votre première expérience d’expérience d’entreprise.

|À faire|À ne pas faire|
|:------|:------|
|Proposer une simple et courte introduction aux actions principales disponibles dans le complément. | Ne pas inclure des informations et des détails qui ne sont pas pertinents pour la prise en main.
|Donner aux utilisateurs la possibilité d’effectuer une action qui aura un impact positif sur leur utilisation du complément. | Ne pas espérer que les utilisateurs découvrent tous les éléments en même temps. Concentrer les efforts sur le type ’action qui fournit le meilleur rendement.
|Créer une expérience utilisateur attrayante que les utilisateurs vont vouloir compléter. | Ne pas forcer les utilisateurs à parcourir toute l’expérience de première utilisation. Donner aux utilisateurs une option leur permettant d’ignorer l’expérience de première exécution. |

Déterminer s’il convient de montrer l’expérience de première utilisation une fois ou plusieurs fois (tout dépend de son importance pour votre scénario). Par exemple, si votre complément est uniquement utilisé de temps en temps, les utilisateurs peuvent devenir moins familiarisés avec le complément. Ils pourraient alors bénéficier d’une autre interaction avec l’expérience de première exécution.

Appliquer les modèles suivants le cas échéant pour créer ou optimisez l’expérience de première exécution de votre complément.

## <a name="carousel"></a>Carrousel

Le carrousel présente aux utilisateurs une série de fonctionnalités ou d’informations avant qu’ils ne commencent à utiliser le complément.

*Figure 1. Autoriser les utilisateurs à faire avancer ou ignorer les pages de début du flux carrousel*

![Illustration montrant l’étape 1 d’un carrousel lors de la première expérience d’utilisation d’un Office d’application de bureau. Dans cet exemple, une action « Ignorer » est incluse dans le haut à droite du volet Des tâches.](../images/add-in-FRE-step-1.png)

*Figure 2. Réduire le nombre d’écrans carrousels uniquement à ce qui est nécessaire pour communiquer efficacement votre message*

![Illustration montrant l’étape 2 d’un carrousel lors de la première expérience d’utilisation d’un Office d’application de bureau. Dans cet exemple, il existe 3 écrans carrousels dans le volet Des tâches.](../images/add-in-FRE-step-2.png)

*Figure 3. Fournir un appel clair à l’action pour quitter l’expérience de première run*

![Illustration montrant l’étape 3 d’un carrousel lors de la première expérience d’utilisation d’un Office d’application de bureau. Dans cet exemple, le troisième et dernier écran du volet Des tâches affiche un bouton pour commencer.](../images/add-in-FRE-step-3.png)

## <a name="value-placemat"></a>Mise en place de la valeur

La mise en place de la valeur communique la proposition de valeur de votre complément en faisant appel au positionnement de votre logo, à une proposition de valeur clairement déclarée, à une présentation ou un résumé des fonctionnalités et à une fonctionnalité claire d’appel à l’action.

*Figure 4. Un lieu de valeur avec logo, proposition de valeur claire, résumé des fonctionnalités et appel à l’action*

![Illustration montrant une valeur de mise en place dans l’expérience de première utilisation d’un Office d’application de bureau. Dans cet exemple, le volet Des tâches affiche le logo du module, une description du module et un bouton pour commencer.](../images/add-in-FRE-value.png)

### <a name="video-placemat"></a>Mise en place vidéo

La mise en place vidéo montre une vidéo aux utilisateurs avant qu’ils ne commencent à utiliser votre complément.

*Figure 5. Première séquence de placemat vidéo - L’écran contient une image fixe de la vidéo avec un bouton lire et effacer le bouton d’appel à l’action*

![Illustration montrant une mise en place vidéo lors de la première expérience d’utilisation d’Office du volet Des tâches de l’application de bureau.](../images/add-in-FRE-video.png)

*Figure 6. Lecteur vidéo : les utilisateurs ont présenté une vidéo dans une fenêtre de boîte de dialogue*

![Illustration montrant une vidéo dans une fenêtre de boîte de dialogue avec une application Office de bureau et le volet Des tâches du add-in en arrière-plan.](../images/add-in-FRE-video-dialog.png)
