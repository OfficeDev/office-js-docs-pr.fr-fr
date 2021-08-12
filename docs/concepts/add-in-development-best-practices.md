---
title: Meilleures pratiques en matière de développement de compléments Office
description: Appliquez les meilleures pratiques lors du développement pour créer des Office de développement.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 2e9a0c3528f9ea6785dc8c2313aa6f9d12db474f2da21b0c8a7cf10c742c23fa
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57083353"
---
# <a name="best-practices-for-developing-office-add-ins"></a>Meilleures pratiques en matière de développement de compléments Office

Des compléments efficaces proposent des fonctionnalités uniques et attrayantes qui étendent les applications Office d’une manière visuellement attractive. Pour créer un complément intéressant, offrez une première expérience attractive à vos utilisateurs, concevez une interface utilisateur de premier choix et optimisez les performances de votre complément. Appliquez les meilleures pratiques décrites dans cet article pour créer des compléments permettant aux utilisateurs d’accomplir leurs tâches rapidement et efficacement.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="provide-clear-value"></a>Indication d’une valeur claire

- Créez des compléments qui aident les utilisateurs à réaliser des tâches rapidement et efficacement. Concentrez-vous sur des scénarios adaptés aux applications Office. Par exemple :
  - Réalisez des tâches de création essentielles plus rapidement et plus facilement, avec moins d’interruptions.
  - Développez de nouveaux scénarios dans Office.
  - Incorporez des services complémentaires dans Office applications.
  - Améliorez l’expérience Office pour accroître la productivité.
- Assurez-vous que la valeur de votre complément apparaîtra clairement aux utilisateurs dès la première utilisation en créant une [première expérience enrichissante](#create-an-engaging-first-run-experience).
- Rédigez une [description claire pour AppSource](/office/dev/store/create-effective-office-store-listings). Soulignez les avantages de votre complément dans votre titre et votre description. Ne comptez pas sur votre marque pour communiquer sur les fonctionnalités de votre complément.

## <a name="create-an-engaging-first-run-experience"></a>Création d’une première expérience intéressante

- Attirez de nouveaux utilisateurs avec une première expérience très simple et intuitive. Les utilisateurs décident toujours d’utiliser ou d’abandonner un complément après l’avoir téléchargé à partir du Windows Store.

- Indiquez clairement les étapes que l’utilisateur doit suivre pour utiliser votre complément. Utilisez des vidéos, des schémas, des panneaux de pagination ou d’autres ressources pour attirer les utilisateurs.

- N’hésitez pas à ajouter un texte pour insister sur l’utilité de votre complément sur l’écran de connexion des utilisateurs.

- Proposez une interface utilisateur pédagogique pour guider les utilisateurs et la personnaliser.

  ![Capture d’écran montrant une comparaison « À faire » et « Ne pas faire ». L’exemple « Faire » montre un add-in qui inclut un bouton sur le bouton sur quoi les utilisateurs peuvent cliquer pour commencer. L’exemple « À ne pas faire » montre un add-in sans étapes ni boutons d’introduction.](../images/contoso-part-catalog-do-dont.png)

- Si votre complément de contenu est lié à des données dans le document de l’utilisateur, incluez des exemples de données ou un modèle pour montrer aux utilisateurs le format de données à utiliser.

  ![Capture d’écran montrant une comparaison « À faire » et « Ne pas faire ». L’exemple « Faire » montre un add-in qui inclut un bouton sur le bouton que les utilisateurs peuvent cliquer pour insérer des exemples de données. L’exemple « À ne pas faire » montre un add-in sans exemple de données ou de boutons.](../images/add-in-title.png)

- Offrez des [essais gratuits](/office/dev/store/decide-on-a-pricing-model). Si votre complément nécessite un abonnement, proposez certaines fonctionnalités gratuitement.

- Facilitez l’inscription. Préremplissez les informations (e-mail, nom d’affichage) et ignorez les vérifications d’adresses e-mail.

- Évitez d’utiliser des fenêtres contextuelles. Si vous devez les utiliser, aidez les utilisateurs à les activer.

Pour les modèles de conception à appliquer lors du développement de votre première expérience d’utilisation, reportez-vous à la section [Modèles de conception de l’expérience utilisateur pour les compléments Office](../design/first-run-experience-patterns.md).

## <a name="use-add-in-commands"></a>Utilisation des commandes de complément

- Fournissez des points d’entrée d’interface utilisateur pertinents pour votre complément à l’aide des commandes de complément. Pour plus d’informations, y compris les bonnes pratiques de conception, reportez-vous aux [commandes de complément](../design/add-in-commands.md).

## <a name="apply-ux-design-principles"></a>Application des principes de conception de l’expérience utilisateur

- Assurez-vous que l’aspect, la convivialité et la fonctionnalité de votre complément améliorent l’expérience Office. Voir [Concevoir l’interface utilisateur de Office de l’interface utilisateur.](../design/add-in-design.md)

- Privilégiez le contenu plutôt que l’apparence. Évitez les éléments d’interface utilisateur superflus qui n’ajoutent pas de valeur à l’expérience utilisateur.

- Gardez le contrôle des utilisateurs. Assurez-vous que ces derniers comprennent les décisions importantes et peuvent facilement rétablir des actions effectuées par le complément.

- Utilisez la personnalisation afin d’inspirer la confiance et d’orienter les utilisateurs. N’utilisez pas la personnalisation afin de submerger les utilisateurs ou de faire de la publicité.

- Évitez d’utiliser le défilement. Optimisez votre complément pour une résolution de 1366 x 768.

- N’incluez pas d’image sans licence.

- Utilisez un [langage clair et simple](../design/voice-guidelines.md) dans votre complément.

- Soulignez l’accessibilité : votre complément doit être facile à utiliser pour tous les utilisateurs et s’accommoder de technologies d’assistance telles que les lecteurs d’écran.

- Adaptez-le à toutes les plateformes et méthodes d’entrée, y compris la souris/le clavier et la [fonction tactile](#optimize-for-touch). Assurez-vous que votre interface utilisateur réagit à différents formats.

### <a name="optimize-for-touch"></a>Optimisation de la fonction tactile

- Utilisez la [propriété Context.touchEnabled](/javascript/api/office/office.context#touchEnabled) pour détecter si l’application Office sur lequel votre application s’exécute est tactile.

  > [!NOTE]
  > Cette propriété n’est pas prise en charge dans Outlook.

- Assurez-vous que toutes les commandes sont correctement dimensionnées pour l’interaction tactile. Par exemple, vérifiez que les boutons disposent de cibles tactiles adéquates et que les zones de texte sont assez grandes pour permettre la saisie.

- N’utilisez pas de méthodes d’entrée non tactiles comme l’utilisation du curseur ou du clic droit.

- Assurez-vous que votre complément fonctionne dans les modes portrait et paysage. Gardez à l’esprit qu’une partie de votre complément pourrait être masquée par le clavier virtuel sur les appareils tactiles.

- Testez votre complément sur un véritable appareil en utilisant le [chargement de version test](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).

> [!NOTE]
> Si vous utilisez des [Fluent’interface](../design/using-office-ui-fabric-react.md) utilisateur React pour vos éléments de conception, la plupart de ces éléments sont intégrés au système de conception.


## <a name="optimize-and-monitor-add-in-performance"></a>Optimisation et contrôle des performances du complément

- Donnez l’impression que l’interface utilisateur réagit rapidement. Votre complément doit se charger en 500 ms au maximum.

- Veillez à ce que toutes les interactions utilisateur répondent en moins d’une seconde.

- Fournissez des indicateurs de chargement pour les opérations à longue durée d’exécution.

- Utilisez un CDN pour héberger les images, les ressources et les bibliothèques communes. Chargez autant d’éléments que possible à partir d’un seul emplacement.

- Suivez les pratiques web standard pour optimiser votre page web. En production, utilisez uniquement les versions réduites des bibliothèques. Chargez uniquement les ressources dont vous avez besoin et optimisez leur chargement.

- Si l’exécution des opérations dure longtemps, fournissez des commentaires aux utilisateurs. Prenez en compte les seuils indiqués dans le tableau suivant. Pour plus d’informations, reportez-vous à l’article sur les [limites des ressources et l’optimisation des performances pour les compléments Office](../concepts/resource-limits-and-performance-optimization.md).

  |Classe d’interaction|Target|Limite supérieure|Perception humaine|
  |:-----|:-----|:-----|:-----|
  |Instantanée|<= 50 ms|100 ms|Aucun délai notable.|
  |Rapide|50-100 ms|200 ms|Délai notable minime. Aucun commentaire n’est nécessaire.|
  |Normale|100-300 ms|500 ms|L’opération va assez vite, sans pour autant pouvoir être qualifiée de rapide. Aucun commentaire n’est nécessaire.|
  |Réactive|300-500 ms|1 seconde|L’opération n’est pas rapide, mais le système donne l’impression de répondre. Aucun commentaire n’est nécessaire.|
  |Continue|> 500 ms|5 secondes|Attente moyenne, le système n’a plus l’air de répondre. Un commentaire peut-être nécessaire.|
  |Captive|> 500 ms|10 secondes|Long, mais pas assez pour faire autre chose. Un commentaire peut-être nécessaire.|
  |Étendue|> 500 ms|> 10 secondes|Assez long pour faire quelque chose en attendant. Un commentaire peut être nécessaire.|
  |Longue durée d’exécution|>5 secondes|>1 minute|Les utilisateurs effectueront certainement une autre action.|

- Surveillez l’état de votre service et utilisez la télémétrie pour surveiller le succès d’utilisateur.

- Réduisez les échanges de données entre le Office document. Pour plus d’informations, voir [Éviter d’utiliser la méthode context.sync en boucles.](correlated-objects-pattern.md)

## <a name="market-your-add-in"></a>Commercialisation de votre complément

- Publiez votre complément dans [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) et [faites sa promotion](/office/dev/store/promote-your-office-store-solution) sur votre site web. Créez un [référencement AppSource efficace](/office/dev/store/create-effective-office-store-listings).

- Utilisez des titres et des descriptifs courts pour le complément. Ils ne doivent pas comporter plus de 128 caractères.

- Rédigez des descriptions brèves et attrayantes pour votre complément. Répondez à la question « Quel problème ce complément résout-il ? ».

- Faites ressortir la proposition de valeur de votre complément dans le titre et la description. Ne comptez pas sur votre marque.

- Créez un site web pour aider les utilisateurs à trouver votre complément et à l’utiliser.

## <a name="use-javascript-that-supports-internet-explorer"></a>Utiliser JavaScript qui prend en charge Internet Explorer

[!INCLUDE [How to support IE](../includes/es5-support.md)]

## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble de la plateforme des compléments Office](../overview/office-add-ins.md)
- [Découvrez le programme pour les développeurs Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
