---
title: Modèles de première expérience d’utilisation des complément Office
description: Découvrez les meilleures pratiques pour la conception d’expériences de première exécution dans des compléments Office.
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: f89656b9c1d1741f38a7122ba11440d2dfca46bf
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608522"
---
# <a name="first-run-experience-patterns"></a>Modèles de première expérience d’utilisation

Une première expérience d’utilisation (FRE) correspond à l’introduction d’un utilisateur à votre complément. Une FRE existe quand un utilisateur ouvre un complément pour la première fois et lui fournit des informations sur les fonctions, les fonctionnalités et/ou les avantages du complément. Cette expérience vous permet de modeler la première impression qu’un utilisateur va avoir d’un complément. Elle peut grandement influencer la probabilité qu’il y revienne et continue à utiliser votre complément...

## <a name="best-practices"></a>Meilleures pratiques


Suivez ces meilleures pratiques lors de la personnalisation de la première expérience d’utilisation :

|À faire|À ne pas faire|
|:------|:------|
|Proposer une simple et courte introduction aux actions principales disponibles dans le complément. | Ne pas inclure des informations et des détails qui ne sont pas pertinents pour la prise en main.
|Donner aux utilisateurs la possibilité d’effectuer une action qui aura un impact positif sur leur utilisation du complément. | Ne pas espérer que les utilisateurs découvrent tous les éléments en même temps. Concentrer les efforts sur le type ’action qui fournit le meilleur rendement.
|Créer une expérience utilisateur attrayante que les utilisateurs vont vouloir compléter. | Ne pas forcer les utilisateurs à parcourir toute l’expérience de première utilisation. Donner aux utilisateurs une option leur permettant d’ignorer l’expérience de première exécution. |



Déterminer s’il convient de montrer l’expérience de première utilisation une fois ou plusieurs fois (tout dépend de son importance pour votre scénario). Par exemple, si votre complément est uniquement utilisé de temps en temps, les utilisateurs peuvent devenir moins familiarisés avec le complément. Ils pourraient alors bénéficier d’une autre interaction avec l’expérience de première exécution.



Appliquer les modèles suivants le cas échéant pour créer ou optimisez l’expérience de première exécution de votre complément.



## <a name="carousel"></a>Carrousel


Le carrousel présente aux utilisateurs une série de fonctionnalités ou d’informations avant qu’ils ne commencent à utiliser le complément.

*Figure 1 : autoriser les utilisateurs à faire avancer ou ignorer les pages d’introduction du flux du carrousel. * 
 ![Première exécution – Carrousel – spécifications pour le volet des tâches](../images/add-in-FRE-step-1.png)



*Figure 2 : Minimiser le nombre d’écrans carrousel que va voir l’utilisateur, pour n’afficher que ce qui est nécessaire pour communiquer efficacement votre message de*
![première exécution – carrousel – spécifications pour le volet de tâches du bureau](../images/add-in-FRE-step-2.png)


*Figure 3 : Fournir un point d’appel bien visible pour permettre à l’utilisateur de quitter l’expérience de première exécution. * 
 ![Première exécution – carrousel – spécifications pour le volet de tâches du bureau](../images/add-in-FRE-step-3.png)



## <a name="value-placemat"></a>Mise en place de la valeur

La mise en place de la valeur communique la proposition de valeur de votre complément en faisant appel au positionnement de votre logo, à une proposition de valeur clairement déclarée, à une présentation ou un résumé des fonctionnalités et à une fonctionnalité claire d’appel à l’action.



![Première exécution – Mise en place de la valeur – spécifications pour le volet des tâches du bureau](../images/add-in-FRE-value.png)
*Une mise en place de la valeur en faisant appel au logo, à une proposition de valeur clairement déclarée, résumé des fonctionnalités et une fonctionnalité claire d’appel à l’action.*


### <a name="video-placemat"></a>Mise en place vidéo

La mise en place vidéo montre une vidéo aux utilisateurs avant qu’ils ne commencent à utiliser votre complément.


*Figure 1 : Mise en place de la première exécution – l’écran contient une image fixe tirées de la vidéo avec un bouton de lecture et ainsi que des boutons d’action clairs. * ![Mise en place vidéo – spécifications pour le volet de tâches du bureau](../images/add-in-FRE-video.png)



*Figure 2 : Lecteur vidéo – on présente aux utilisateurs une vidéo incluse dans une boite de dialogue. *
![Mise en place vidéo – spécifications pour le volet de tâches du bureau](../images/add-in-FRE-video-dialog.png)
