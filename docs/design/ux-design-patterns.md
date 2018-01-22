# <a name="ux-design-pattern-templates-for-office-add-ins"></a>Modèles de conception d’expérience utilisateur pour les compléments Office 

Le [projet de modèles de conception de l’expérience utilisateur pour compléments Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code "projet de modèles de conception de l’expérience utilisateur pour compléments Office") inclut des fichiers HTML, JavaScript et CSS que vous pouvez utiliser pour créer l’expérience utilisateur de votre complément.   

Utiliser le projet de modèles de conception d’expérience utilisateur aux fins suivantes :

* Appliquer des solutions à des scénarios client courants.
* Appliquer les meilleures pratiques en matière de conception.
* Incorporer les composants et styles d’[Office UI Fabric](https://dev.office.com/fabric#/get-started).
* Créer des compléments qui s’intègrent visuellement à l’interface utilisateur d’Office par défaut.  

## <a name="using-the-ux-design-patterns"></a>Utilisation des modèles de conception d’expérience utilisateur

Vous pouvez vous guider à l’aide des [spécifications des modèles de conception d’expérience utilisateur](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns) lorsque vous créez votre propre complément Office. Vous pouvez également ajouter le [code source](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates) directement dans votre projet.

Pour utiliser les spécifications afin de créer une maquette de votre propre interface utilisateur du complément, procédez comme suit :

1. Téléchargez les fichiers de ressources de conception et commencez à concevoir votre propre interface utilisateur :
    * [Composants de conception d’expérience utilisateur de complément Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/addin_ux_design_components.ai) (fichier Adobe Illustrator)
    * [Modèles de conception d’expérience utilisateur de complément Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/addin_ux_design_patterns.ai) (fichier Adobe Illustrator)
    * [Prototype de conception d’expérience utilisateur de complément Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/addin_ux_design_prototype.xd) (fichier Adobe Experience Design)
2. Pour obtenir des instructions, reportez-vous aux articles suivants :
    * [Modèles de conception d’expérience utilisateur](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/README.md)
    * Bonnes pratiques en matière de [conception de compléments Office](https://dev.office.com/docs/add-ins/design/add-in-design)
    * [Kits d’outils Office UI Fabric](https://developer.microsoft.com/fr-FR/fabric#/resources)

Pour ajouter le code source, procédez comme suit :

1. Clonez le [référentiel du projet de modèles de conception de l’expérience utilisateur pour les compléments Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code "projet de modèles de conception de l’expérience utilisateur pour les compléments Office"). 
2. Copiez le [dossier des composants](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/assets) ainsi que le dossier de code pour le modèle individuel que vous choisissez dans votre projet de complément.  
3. Incorporez le modèle individuel à votre complément. Par exemple :
    - Modifiez l’emplacement source ou l’URL de commande de complément dans le manifeste.
    - Utilisez le modèle de conception d’expérience utilisateur en tant que modèle pour d’autres pages.
    - Lien vers ou à partir du modèle de conception d’expérience utilisateur.

> **Remarque :** certaines spécifications de modèle d’expérience utilisateur ne correspondent pas au code source. Nous mettons tout en œuvre pour aligner toutes les ressources. Notez également que certaines spécifications sont présentées comme archivées. Nous évaluons la valeur de ces spécifications archivées sur la plateforme. Chaque modèle vise à représenter un modèle unique et d’interaction. Les modèles ne doivent pas se chevaucher et doivent être bien différenciés des composants Office Fabric UI.

## <a name="types-of-ux-design-patterns"></a>Types de modèles de conception de l’expérience utilisateur
### <a name="generic-pages"></a>Pages génériques

Les modèles de page générique peuvent être appliqués à n’importe quelle page de votre complément et n’ont pas d’usage particulier. L’un des modèles de première utilisation constitue un exemple de page à usage spécifique. La liste suivante décrit les pages génériques disponibles :

* **Page d’accueil** : une page de complément standard, par exemple la page sur laquelle un utilisateur est renvoyé après une première expérience d’utilisation ou un processus de connexion. 
    * En savoir plus sur les instructions relatives à l’adoption du [langage de conception Office](https://dev.office.com/docs/add-ins/design/add-in-design-language) dans votre complément.
    * [Code de la page d’accueil](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/landing-page)
* **Image de marque dans la barre de marque** - La page d’accueil avec une image dans le pied de page qui représente votre marque. 
    * [Spécification de la barre de marque](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/brand-bar.md)
    * [Code de la barre de marque](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/brand-bar)

<table>
 <tr><th>Accueil</th><th>Barre de marque</th></tr>
 <tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/landing-page"><img src="../images/landing.page.PNG" alt="landing page" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/brand-bar"><img src="../images/brand.bar.PNG" alt="brand bar" style="width: 264px;"/></A></td></tr>
 </table>
 
### <a name="first-run-experience"></a>Première expérience d’utilisation

Il s’agit de l’expérience vécue par un utilisateur lorsqu’il ouvre votre complément pour la première fois. Les modèles de modèle de conception de première utilisation suivants sont disponibles : 

* **Étapes de démarrage** - Permet aux utilisateurs ayant une liste d’étapes à suivre de commencer à utiliser votre complément. 
    * [Spécification des étapes de démarrage](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/fre_stepsToStart.pdf)
        * Ce modèle de conception d’expérience utilisateur a été archivé. Sachant que nous évaluons sa valeur, reportez-vous également à la [spécification sur la valeur de la première exécution](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/value-placemat.md).  
    * [Code des étapes de démarrage](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/instruction-step)
* **Valeur** - Communique la proposition de valeur de votre complément.
    * [Spécification de la valeur](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/value-placemat.md)
    * [Code de la valeur](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/value-placemat)
* **Vidéo** - Montre une vidéo aux utilisateurs avant qu’ils commencent à utiliser votre complément.
    * [Spécification de la vidéo](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/video-placemat.md)
    * [Code de la vidéo](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat)
* **Procédure pas à pas** : explique aux utilisateurs une série de fonctionnalités ou d’informations avant qu’ils commencent à utiliser le complément.
    * [Spécification de Carrousel](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/carousel.md)
        * Remarque : ce modèle de conception d’expérience utilisateur a été renommé « Carrousel ». Les anciennes spécifications y faisaient référence comme « Panneau de pagination ». Les ressources de code y font référence comme « Procédure pas à pas pour la première exécution ». 
    * [Code de la procédure pas à pas](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/walkthrough)

L’[Office Store](https://msdn.microsoft.com/fr-FR/library/office/jj220033.aspx) dispose d’un système qui gère les versions d’évaluation d’un complément, mais si vous souhaitez contrôler l’interface utilisateur relative à l’expérience d’évaluation de votre complément, utilisez les modèles suivants :

* **Version d’évaluation** - Explique aux utilisateurs comment utiliser la version d’évaluation de votre complément.
    * [Spécification de la version d’évaluation](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/fre_trialVersion.pdf)
        * Ce modèle de conception d’expérience utilisateur a été archivé. Sachant que nous évaluons sa valeur, veuillez vous reporter au fichier PDF.
    * [Code de la version d’évaluation](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat)
* **Fonctionnalité d’évaluation** - Informe les utilisateurs que la fonctionnalité qu’ils tentent d’utiliser n’est pas disponible dans la version d’évaluation du complément. Par ailleurs, si votre complément est gratuit, mais qu’il comporte une fonctionnalité qui nécessite un abonnement, envisagez d’utiliser ce modèle. Vous pouvez également utiliser ce modèle pour offrir une expérience avec une version antérieure après qu’une période d’évaluation est terminée.
    * [Spécification de la fonctionnalité d’évaluation](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/fre_trialFeature.pdf)
        * Ce modèle de conception d’expérience utilisateur a été archivé. Sachant que nous évaluons sa valeur, veuillez vous reporter au PDF ci-dessus.
    * [Code de la fonctionnalité d’évaluation](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat-feature)

> **Important :** Si vous décidez de gérer votre propre version d’évaluation et de ne pas utiliser l’Office Store pour gérer la version d’évaluation, assurez-vous que vous incluez la balise **Un autre achat peut être requis** dans les notes de test du service Mon tableau de bord vendeur.

Déterminez s’il convient de montrer la vidéo sur la première expérience d’utilisation une ou plusieurs fois (tout dépend de son importance pour votre scénario). Par exemple, si les utilisateurs utilisent votre complément régulièrement, ils peuvent oublier comment l’utiliser. Il peut être utile de consulter la première expérience d’utilisation plusieurs fois. 

 <table>
 <tr><th>Étapes de démarrage</th><th>Valeur</th><th>Vidéo</th></tr>
 <tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/instruction-step"><img src="../images/instruction.step.PNG" alt="instruction steps" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/value-placemat"><img src="../images/value.placemat.PNG" alt="value placemat" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat"><img src="../images/video.placemat.PNG" alt="video placemat" style="width: 264px;"/></A></td></tr>
 </table>

 <table>
 <tr><th>Première page de la procédure pas à pas</th><th>Version d’évaluation</th><th>Fonctionnalité d’évaluation</th></tr>
 <tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/walkthrough"><img src="../images/walkthrough1.PNG" alt="walkthrough 1" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat"><img src="../images/trial.placemat.PNG" alt="trial placemat" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat-feature"><img src="../images/trial.placemat.feature.PNG" alt="trial placemat feature" style="width: 264px;"/></A></td></tr>
 </table> 

### <a name="navigation"></a>Navigation

Les utilisateurs doivent naviguer entre les différentes pages de votre complément. Les modèles de navigation suivants indiquent différentes options que vous pouvez utiliser afin d’organiser les pages et les commandes de votre complément.

* **Bouton Page précédente et Page suivante** : affiche un volet Office avec les boutons Page précédente et Page suivante. Utilisez ce modèle pour vous assurer que les utilisateurs suivent une série d’étapes ordonnées.
    * [Spécification des boutons Page précédente et Page suivante](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/back-button.md)
    * [Code des boutons Page précédente et Page suivante](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/back-button) 
* **Navigation** - Affiche un menu, communément appelé menu hamburger, avec les éléments de menu de la page dans un volet Office. 
    * [Spécification de la navigation](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/contextual-menu.md)
    * [Code de la navigation](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation) 
* **Navigation à l’aide de commandes** - Affiche le menu hamburger avec les boutons de commande (ou d’action) dans un volet Office. Utilisez ce modèle lorsque vous voulez fournir des options de navigation et de commande ensemble. 
    * [Spécification de la navigation à l’aide de commandes](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/command-bar.md)
    * [Code de la navigation à l’aide de commandes](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation-commands)
* **Tableau croisé dynamique** - Affiche la navigation du tableau croisé dynamique dans un volet Office. Utilisez la navigation du tableau croisé dynamique pour permettre aux utilisateurs de naviguer entre les différents contenus.
    * [Spécification du tableau croisé dynamique](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/pivot.md)
    * [Code du tableau croisé dynamique](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/pivot)
* **Barre d’onglets** - Affiche la navigation à l’aide de boutons avec du texte et des icônes verticalement empilés. Utiliser la barre d’onglets pour permettre la navigation à l’aide des onglets avec des titres courts et explicites.
    * [Spécification de la barre d’onglets](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/tab-bar.md)
    * [Code de la barre d’onglets](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/tab-bar) 

<table>
<tr><th>Bouton Précédent</th><th>Navigation</th><th>Navigation à l’aide de commandes</th></tr>
<tr>
    <td>
        <A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/back-button">
        <img src="../images/back.button.png" alt="back button" style="width: 264px;"/></A>
    </td>
    <td>
        <A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation">
        <img src="../images/navigation.png" alt="navigation" style="width: 264px;"/></A>
    </td>
    <td>
        <A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation-commands">
        <img src="../images/navigation.commands.png" alt="navigation with commands" style="width: 264px;"/></A>
    </td>
</tr>
 </table>

<table>
<tr><th>Pivot</th><th>Barre d’onglets</th></tr>
<tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/pivot">
<img src="../images/pivot.png" alt="pivot navigation" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/tab-bar">
<img src="../images/tab.bar.png" alt="tab bar" style="width: 264px;"/></A></td>
</tr>
 </table>

### <a name="notifications"></a>Notifications

Votre complément peut avertir les utilisateurs d’événements, tels qu’une erreur, ou de l’état d’avancement d’un élément de plusieurs façons. Les modèles de notification suivants sont disponibles : 

* **Boîte de dialogue incorporée** - Affiche une boîte de dialogue dans le volet des tâches qui vous fournit des informations et, éventuellement, une expérience interactive, à l’aide des boutons ou d’autres commandes. Pensez à en utiliser une pour inviter un utilisateur à confirmer une action. Utiliser le modèle de boîte de dialogue incorporée lorsque vous souhaitez conserver l’expérience utilisateur dans le volet Office.
    * [Spécification de la boîte de dialogue incorporée](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/embedded-dialog.md)
    * [Code de la boîte de dialogue incorporée](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/embedded-dialog)
* **Message incorporé** - Indique l’échec, la réussite ou des informations, et peut apparaître à un emplacement spécifié dans le volet Office. Par exemple, si un utilisateur entre une adresse de messagerie incorrecte dans une zone de texte, un message d’erreur apparaît juste en dessous de la zone de texte. 
    * [Spécification du message incorporé](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/notification_inlineMessage.pdf)
        * Ce modèle de conception d’expérience utilisateur a été archivé. Sachant que nous évaluons sa valeur, veuillez vous reporter au PDF ci-dessus.
    * [Code du message incorporé](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/inline-message)
* **Bannière de message** - Fournit des informations et, éventuellement, des instructions dans une bannière qui peut être réduite à une seule ligne, développée en plusieurs lignes ou masquée. Utilisez des bannières de message pour signaler une mise à jour du service ou donner un conseil utile lorsque le complément démarre. 
    * [Spécification de la bannière de message](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/message_bar.pdf)
        * Ce modèle de conception d’expérience utilisateur a été archivé. Sachant que nous évaluons sa valeur, veuillez vous reporter au PDF ci-dessus.
    * [Code de la bannière de message](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/message-banner)
* **Barre de progression** - Indique la progression d’un processus long et synchrone, tel qu’une tâche de configuration qui doit être terminée pour que l’utilisateur puisse effectuer d’autres actions. Il s’agit d’une page distincte interstitielle qui met en évidence la marque du complément. Utilisez une barre de progression quand le processus peut envoyer des notifications pour indiquer la progression de la tâche dans le complément.
    * [Spécification de la barre de progression](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/progress-indicator.md)
    * [Code de la barre de progression](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/progress-bar)
* **Bouton fléché** - Indique qu’un processus synchrone long est lancé, mais ne fournit aucune indication sur son état d’avancement. Il s’agit d’une page distincte interstitielle qui met en évidence la marque du complément. Utilisez un bouton fléché quand le complément ne peut pas indiquer avec précision la progression du processus. 
    * [Spécification du bouton fléché](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/spinner.md)
    * [Code du bouton fléché](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/spinner)
* **Annonce** - Fournit un bref message qui disparaît au bout de quelques secondes. Comme il se peut que l’utilisateur ne voie pas le message, utilisez une annonce uniquement pour les informations non importantes. Utilisez une annonce pour informer les utilisateurs d’un événement dans un système distant, tel que la réception d’un message électronique.
    * [Spécification de l’annonce](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/toast.md)
    * [Code de l’annonce](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/toast)

 <table>
 <tr><th>Boîte de dialogue incorporée</th><th>Message incorporé</th><th>Bannière de message</th></tr>
 <tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/embedded-dialog"><img src="../images/embedded.dialog.PNG" alt="embedded dialog" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/inline-message"><img src="../images/inline.message.PNG" alt="inline message" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/message-banner"><img src="../images/message.banner.PNG" alt="message banner" style="width: 264px;"/></A></td></tr>
 </table>

 <table>
 <tr><th>Barre de progression</th><th>Bouton fléché</th><th>Annonce</th></tr>
 <tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/progress-bar"><img src="../images/progress.bar.PNG" alt="progress bar" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/spinner"><img src="../images/spinner.PNG" alt="spinner" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/toast"><img src="../images/toast.PNG" alt="toast" style="width: 264px;"/></A></td></tr>
 </table>
 


### <a name="general-components"></a>Composants généraux

Les éléments suivants constituent des composants généraux que vous pouvez utiliser avec vos compléments dans différents scénarios.  

#### <a name="client-dialog-boxes"></a>Boîtes de dialogue client

Les boîtes de dialogue client fournissent aux utilisateurs un autre moyen de travailler avec votre complément en dehors d’un volet Office. Les modèles de boîte de dialogue suivants sont disponibles :

* **Boîte de dialogue de rampe de type** : affiche une boîte de dialogue avec du contenu textuel. Utilisez la boîte de dialogue de rampe de type pour transmettre des informations détaillées aux utilisateurs. 
    * En savoir plus sur la conception de [boîtes de dialogue dans les compléments Office](https://dev.office.com/docs/add-ins/design/dialog-boxes). Suivez également nos recommandations concernant la [typographie dans les compléments Office](https://dev.office.com/docs/add-ins/design/add-in-design-language#typography).
    * [Code de la boîte de dialogue de rampe de type](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/typeramp)
* **Boîte de dialogue d’alerte** - Affiche un message d’alerte avec des informations importantes, comme les erreurs ou les notifications, aux utilisateurs.  
    * [Spécification de la boîte de dialogue d’alerte](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/notification_alert.pdf)
        * Ce modèle de conception d’expérience utilisateur a été archivé. Sachant que nous évaluons sa valeur, veuillez vous reporter au PDF ci-dessus.
    * [Code de la boîte de dialogue d’alerte](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/alert)
* **Boîte de dialogue de navigation** - Affiche une boîte de dialogue comportant la navigation. Utilisez la boîte de dialogue de navigation pour permettre aux utilisateurs de naviguer entre les différents contenus. 
    * En savoir plus sur la conception de [boîtes de dialogue dans les compléments Office](https://dev.office.com/docs/add-ins/design/dialog-boxes). Découvrez également comment utiliser les [composants de tableau croisé dynamique Office UI Fabric dans les compléments Office](https://dev.office.com/docs/add-ins/design/pivot).
    * [Code de la boîte de dialogue de navigation](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation)

<table>
 <tr><th>Boîte de dialogue de rampe de type</th><th>Boîte de dialogue d’alerte</th></tr>
<tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/typeramp"><img src="../images/typeramp.dialog.png" alt="typeramp dialog" style="width: 300px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/alert"><img src="../images/alert.dialog.png" alt="alert dialog" style="width: 264px;"/></A></td>
</tr></tr>
 </table>
 
 <table>
 <tr><th>Boîte de dialogue de navigation</th></tr>
<tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation"><img src="../images/navigation.dialog.png" alt="navigation dialog" style="width: 300px;"/></A></td></tr>
</tr>
 </table>


#### <a name="feedback-and-ratings"></a>Évaluations et commentaires

Pour améliorer la visibilité et l’adoption de votre complément, il est utile de fournir aux utilisateurs la possibilité de noter et de commenter votre complément dans l’Office Store. Ce modèle comporte deux méthodes pour effectuer des commentaires et des évaluations dans le complément :

- Commentaires initiés par l’utilisateur - Un utilisateur choisit d’envoyer des commentaires à l’aide du menu de navigation (par exemple, en utilisant le lien **Envoyer des commentaires**) ou d’une icône dans le pied de page.
- Commentaires initiés par le système - Une fois le complément exécuté trois fois, l’utilisateur est invité à fournir un commentaire, via une bannière de message.

Les deux méthodes ouvrent une boîte de dialogue qui contient la page de l’Office Store pour le complément.

* [Spécification des évaluations et commentaires](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/notification_feedback.pdf)
    * Ce modèle de conception d’expérience utilisateur a été archivé. Sachant que nous évaluons sa valeur, veuillez vous reporter au PDF ci-dessus.
* [Code des évaluations et commentaires](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/feedback/office-store)

>**Important :** Ce modèle pointe actuellement vers la page d’accueil de l’Office Store. Veillez à mettre à jour l’URL avec l’URL de la page de votre complément dans l’Office Store.

 <table>
 <tr><th>Évaluations et commentaires</th></tr>
<tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/feedback/office-store"><img src="../images/feedback.ratings.PNG" alt="Feedback and Ratings" style="width: 264px;"/></A></td></tr>
</tr>
 </table>

#### <a name="settings-and-privacy"></a>Paramètres et confidentialité

Les compléments peuvent nécessiter une page des paramètres afin que les utilisateurs puissent configurer les paramètres qui contrôlent le comportement du complément. Vous pouvez également fournir aux utilisateurs les politiques de confidentialité auxquelles votre complément adhère. 

* **Paramètres** : affiche un volet Office avec des composants de configuration contrôlant le comportement du complément. Une page des paramètres fournit des options que l’utilisateur peut choisir.
    * [Spécification des paramètres](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/settings.md)
    * [Code des paramètres](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings)
* **Politique de confidentialité** - Affiche un volet Office contenant des informations importantes sur les politiques de confidentialité. 
    * [Spécification de la politique de confidentialité](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/general_multiSection.pdf)
        * Ce modèle de conception d’expérience utilisateur a été archivé. Sachant que nous évaluons sa valeur, veuillez vous reporter au PDF ci-dessus.
    * [Code de la politique de confidentialité](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings)

<table>
 <tr><th>Paramètres</th><th>Politique de confidentialité</th></tr>
<tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings"><img src="../images/settings.png" alt="settings" style="width: 300px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings"><img src="../images/privacy.policy.png" alt="privacy" style="width: 264px;"/></A></td>
</tr></tr>
 </table>

## <a name="additional-resources"></a>Ressources supplémentaires

* [Meilleures pratiques en matière de développement de compléments Office](https://dev.office.com/docs/add-ins/overview/add-in-development-best-practices)
* [Office UI Fabric](http://dev.office.com/fabric/)
