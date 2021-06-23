---
title: Conception des compléments Outlook
description: Les instructions suivantes vous aideront à concevoir et à créer un complément attrayant, qui apportera le meilleur de votre application directement dans Outlook sur Windows, le web, iOS, Mac et Android.
ms.date: 06/24/2019
localization_priority: Priority
ms.openlocfilehash: a669d2cf0a98ffa0ca7b7dfc3fcc5b71d291a0e0
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53077133"
---
# <a name="outlook-add-in-design-guidelines"></a>Instructions de création d’un complément Outlook

Les compléments sont un excellent moyen pour les partenaires d’étendre les fonctionnalités d’Outlook au-delà de notre ensemble de fonctionnalités de base. Les compléments permettent aux utilisateurs d’accéder à des expériences, des tâches et du contenu de tiers sans avoir à quitter leur boîte de réception. Une fois installés, les compléments Outlook sont disponibles sur toutes les plateformes et tous les appareils.  

Les instructions de haut niveau suivantes vous aideront à concevoir et à créer un complément attrayant, qui apportera le meilleur de votre application directement dans Outlook&mdash; sur Windows, le web, iOS, Mac et Android.

## <a name="principles"></a>Principes

1. **Concentrez-vous sur quelques tâches clés et exécutez-les correctement**

   Les compléments les mieux conçus sont simples à utiliser, visent un objectif précis et sont réellement utiles pour les utilisateurs. Votre complément s’exécutera dans Outlook, ce principe est donc d’autant plus important. Outlook est une application de productivité&mdash;c’est l’endroit où les utilisateurs se rendent pour s’acquitter de leurs tâches.

   Vous allez apporter une extension à notre expérience et vous devez être certain que les scénarios que vous activez s’intègre naturellement au sein d’Outlook. Réfléchissez bien aux situations dans lesquelles la présence des compléments sera le plus utile pour les utilisateurs dans les expériences de messagerie et de calendrier.

   Un complément ne doit pas tenter d’exécuter tout ce que votre application fait déjà. Concentrez-vous sur les actions appropriées les plus fréquemment utilisées, dans le contexte de contenu Outlook. Pensez à votre appel à l’action et indiquez clairement à l’utilisateur ce qu’il doit faire lorsque votre volet de tâches s’ouvre.

2. **Faites en sorte que tout semble aussi naturel que possible**

   Votre complément doit être conçu à l’aide de schémas natifs de la plateforme sur laquelle Outlook s’exécute. Pour ce faire, veillez à respecter et implémenter les instructions d’interaction et visuelles définies par chaque plateforme. Outlook possède ses propres instructions et celles-ci doivent également être prises en compte. Un complément bien conçu sera une combinaison appropriée de votre expérience, de la plateforme et d’Outlook.

   Cela ne signifie pas que votre complément devra être différent visuellement lorsqu’il est exécuté sur Outlook sur iOS et Outlook sur Android. Nous vous recommandons de vous référer à [Framework7](https://framework7.io/) comme une option pour vous aider dans les styles.

3. **Faites en sorte que votre complément soit agréable à utiliser jusque dans les moindres détails**

   Les gens apprécient les produits qui sont attrayants sur le plan fonctionnel et visuel. Vous pouvez contribuer à garantir le succès de votre complément en créant une expérience où vous avez tenu soigneusement compte de chaque interaction et détail visuel. Les étapes nécessaires à l’exécution d’une tâche doivent être claires et pertinentes. Dans l’idéal, aucune action ne doit exiger plus d’un clic ou deux. 
   
   Un utilisateur ne doit pas sortir du contexte pertinent pour effectuer une action. Un utilisateur doit pouvoir accéder à votre complément et en sortir facilement pour revenir à ce qu’il faisait avant. Un complément n’est pas destiné à être un emplacement où l’utilisateur passe beaucoup de temps&mdash;il doit s’agir d’une amélioration de nos fonctionnalités principales. Si votre complément est développé correctement, il nous aidera à augmenter la productivité des utilisateurs, ce qui constitue un de nos objectifs.

4. **Personnalisez votre complément à l’image de votre marque de manière judicieuse**

   Nous apprécions les personnalisations et nous savons qu’il est important pour vous de procurer votre expérience unique aux utilisateurs. Cependant, nous pensons que la meilleure façon de garantir la réussite de votre complément est de créer une expérience intuitive qui incorpore subtilement les éléments de votre marque au lieu d’afficher des éléments de marque permanents ou imposants qui empêchent les utilisateurs de naviguer dans votre système de manière fluide. 
    
   Vous pouvez par exemple intégrer votre marque en utilisant les couleurs, les icônes et le ton qui la définissent&mdash;tout en respectant les modèles privilégiés de la plateforme et les critères d’accessibilité. Efforcez-vous de toujours privilégier le contenu et la capacité à effectuer des tâches plutôt que de chercher à attirer l’attention sur votre marque. 
    
   > [!NOTE]
   >  Les publicités ne doivent pas être affichées dans des compléments sur iOS ou Android.

## <a name="design-patterns"></a>Modèles de conception

> [!NOTE]
> Tandis que les principes ci-dessus s’appliquent à l’ensemble des points de terminaison/plateformes, les modèles et les exemples suivants sont spécifiques des compléments mobiles sur la plateforme iOS.

Pour vous aider à créer un complément bien conçu, nous proposons des [modèles](../design/ux-design-pattern-templates.md) pour les versions mobiles avec iOS fonctionnant dans l’environnement Outlook Mobile. Si vous utilisez ces modèles spécifiques, votre complément semblera natif de la plateforme iOS et d’Outlook Mobile. Ces modèles sont également décrits en détail ci-dessous. Bien que cette bibliothèque ne soit pas exhaustive, il s’agit du début de son développement et nous continuerons à l’enrichir à mesure que nous découvrirons des paradigmes que nos partenaires souhaitent inclure dans leurs compléments.  

### <a name="overview"></a>Vue d’ensemble

Un complément type est constitué des éléments suivants.

![Diagramme de modèles d’expérience utilisateur de base pour un volet de tâches sur iOS.](../images/outlook-mobile-design-overview.png)

![Diagramme de modèles d’expérience utilisateur de base pour un volet de tâches sur Android.](../images/outlook-mobile-design-overview-android.jpg)

### <a name="loading"></a>Chargement

Lorsqu’un utilisateur sélectionne votre complément, l’expérience utilisateur doit s’afficher rapidement. Si le chargement est long, utilisez une barre de progression ou un indicateur d’activité. Une barre de progression doit être utilisée lorsque le délai peut être déterminé et un indicateur d’activité doit être utilisé lorsque le délai ne peut pas être déterminé.

**Exemple de chargement de pages sur iOS**

![Exemples illustrant une barre de progression et un indicateur d’activité sur iOS.](../images/outlook-mobile-design-loading.png)

**Exemple de chargement de pages sur Android**

![Exemples illustrant une barre de progression et un indicateur d’activité sur Android.](../images/outlook-mobile-design-loading-android.jpg)


### <a name="sign-insign-up"></a>Connexion/Inscription

Votre procédure de connexion (et d’inscription) doit être directe et simple.

**Exemple de page de connexion et d’inscription sur iOS**

![Exemples de pages pour se connecter et s’inscrire sur iOS.](../images/outlook-mobile-design-signin.png)

**Exemple de page de connexion sur Android**

![Exemples de page pour se connecter sur Android.](../images/outlook-mobile-design-signin-android.png)

### <a name="brand-bar"></a>Barre de marque

Le premier écran de votre complément doit inclure un élément de votre marque. Conçue pour que votre marque soit reconnue, la barre de marque vous aide également à définir le contexte pour l’utilisateur. Étant donné que la barre de navigation contient le nom de votre société/marque, il est inutile de reproduire la barre de marque sur les pages suivantes.

**Exemple de personnalisation sur iOS**

![Exemples de barres de marque sur iOS.](../images/outlook-mobile-design-branding.png)

**Exemple de personnalisation sur Android**

![Exemples de barres de marque sur Android.](../images/outlook-mobile-design-branding-android.png)

### <a name="margins"></a>Marges

Les marges sur mobile doivent être définies sur 15 px (8 % de l’écran) pour chaque côté afin de s’aligner sur Outlook iOS et sur 16 px pour chaque côté afin de s’aligner sur Outlook Android.

![Exemples de marges sur iOS.](../images/outlook-mobile-design-margins.png)

### <a name="typography"></a>Typographie

La typographie est alignée sur Outlook iOS et doit être simple pour la lisibilité.

**Typographie sur iOS**

![Exemples de typographie pour iOS.](../images/outlook-mobile-design-typography.png)

**Typographie sur Android**

![Exemples de typographie pour Android.](../images/outlook-mobile-design-typography-android.png)

### <a name="color-palette"></a>Palette de couleurs

L’utilisation des couleurs est subtile dans Outlook iOS.  À des fins de cohérence, nous vous demandons d’utiliser les couleurs uniquement sur les actions et les erreurs, et que seule la barre de marque utilise une couleur unique.

![Palette de couleurs pour iOS.](../images/outlook-mobile-design-color-palette.png)

### <a name="cells"></a>Cellules

Étant donné que la barre de navigation ne peut pas être utilisée pour libeller une page, utilisez les titres de section pour libeller les pages.

**Exemples de cellules sur iOS**

![Types de cellules pour iOS.](../images/outlook-mobile-design-cell-types.png)
* * *
![Cellules « Do » pour iOS.](../images/outlook-mobile-design-cell-dos.png)
* * *
![Cellules « Don’t » pour iOS.](../images/outlook-mobile-design-cell-donts.png)
* * *
![Cellules et entrées pour iOS.](../images/outlook-mobile-design-cell-input.png)

**Exemples de cellules sur Android**

![Types de cellules pour Android.](../images/outlook-mobile-design-cell-type-android.png)
* * *
![Cellules « Do » pour Android.](../images/outlook-mobile-design-cell-dos-android.png)
* * *
![Cellules « Don’t » pour Android.](../images/outlook-mobile-design-cell-donts-android.png)
* * *
![Cellules et entrées pour Android, partie 1.](../images/outlook-mobile-design-cell-input-1-android.png)

![Cellules et entrées pour Android, partie 2.](../images/outlook-mobile-design-cell-input-2-android.png)

### <a name="actions"></a>Actions

Même si votre application gère une multitude d’actions, réfléchissez aux plus importantes que vous souhaitez que votre complément effectue, et concentrez-vous sur celles-ci.

**Exemples d’actions sur iOS**

![Actions et cellules dans iOS.](../images/outlook-mobile-design-action-cells.png)
* * *
![Actions « Do » pour iOS.](../images/outlook-mobile-design-action-dos.png)

**Exemples d’actions sur Android**

![Actions et cellules dans Android.](../images/outlook-mobile-design-action-cells-android.png)
* * *
![Actions « Do » pour Android.](../images/outlook-mobile-design-action-dos-android.png)

### <a name="buttons"></a>Boutons

Les boutons sont utilisés lorsqu’il existe d’autres éléments de l’expérience utilisateur en dessous (par opposition aux actions, car une action est toujours le dernier élément de l’écran).

**Exemples de boutons sur iOS**

![Exemples de boutons pour iOS.](../images/outlook-mobile-design-buttons.png)

**Exemples de boutons sur Android**

![Exemples de boutons pour Android.](../images/outlook-mobile-design-buttons-android.png)

### <a name="tabs"></a>Onglets

Les onglets peuvent contribuer à organiser le contenu.

**Exemples d’onglets sur iOS**

![Exemples d’onglets pour iOS.](../images/outlook-mobile-design-tabs.png)

**Exemples d’onglets sur Android**

![Exemples d’onglets pour Android.](../images/outlook-mobile-design-tabs-android.png)

### <a name="icons"></a>Icônes

Les icônes doivent respecter la conception Outlook iOS actuelle autant que possible. Utilisez la taille et la couleur standard.

**Exemples d’icônes sur iOS**

![Exemples d’icônes pour iOS.](../images/outlook-mobile-design-icons.png)

**Exemples d’icônes sur Android**

![Exemples d’icônes pour Android.](../images/outlook-mobile-design-icons-android.jpg)

## <a name="end-to-end-examples"></a>Exemples de bout en bout

Pour le lancement de nos compléments Outlook Mobile v1, nous avons travaillé en étroite collaboration avec nos partenaires qui créaient des compléments. Pour présenter le potentiel de leurs compléments sur Outlook Mobile, notre concepteur a regroupé des flux de bout en bout pour chaque complément, en respectant nos instructions et en utilisant nos modèles.

> [!IMPORTANT]
> Ces exemples sont destinés à mettre en évidence la façon idéale de combiner interaction et conception visuelle pour un complément et peuvent ne pas correspondre aux ensembles de fonctionnalités exacts des compléments réels. 

### <a name="giphy"></a>GIPHY

**Exemple de GIPHY sur iOS**

![Conception de bout en bout pour le complément GIPHY sur iOS.](../images/outlook-mobile-design-giphy.png)

**Exemple de GIPHY sur Android**

![Conception de bout en bout pour le complément GIPHY sur Android.](../images/outlook-mobile-design-giphy-android.png)

### <a name="nimble"></a>Nimble

**Exemple de Nimble sur iOS**

![Conception de bout en bout pour le complément Nimble sur iOS.](../images/outlook-mobile-design-nimble.png)

**Exemple de Nimble sur Android**

![Conception de bout en bout pour le complément Nimble sur Android.](../images/outlook-mobile-design-nimble-android.png)

### <a name="trello"></a>Trello

**Exemple de Trello sur iOS**

![Conception de bout en bout pour le complément Trello partie 1 sur iOS.](../images/outlook-mobile-design-trello-1.png)
* * *
![Conception de bout en bout pour le complément Trello partie 2 sur iOS.](../images/outlook-mobile-design-trello-2.png)
* * *
![Conception de bout en bout pour le complément Trello partie 3 sur iOS.](../images/outlook-mobile-design-trello-3.png)

**Exemple de Trello sur Android**

![Conception de bout en bout pour le complément Trello partie 1 sur Android.](../images/outlook-mobile-design-trello-1-android.png)
* * *
![Conception de bout en bout pour le complément Trello partie 2 sur Android.](../images/outlook-mobile-design-trello-2-android.png)

### <a name="dynamics-crm"></a>Dynamics CRM

**Exemple de Dynamics CRM sur iOS**

![Conception de bout en bout pour le complément Dynamics CRM sur iOS.](../images/outlook-mobile-design-crm.png)

**Exemple de Dynamics CRM sur Android**

![Conception de bout en bout pour le complément Dynamics CRM sur Android.](../images/outlook-mobile-design-crm-android.png)
