---
title: Mod?les de conception d?exp?rience utilisateur pour les compl?ments Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: c8ec23db5e7c4c571babff94bdc617b78340d965
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="ux-design-pattern-templates-for-office-add-ins"></a>Mod?les de conception d?exp?rience utilisateur pour les compl?ments Office

Le [projet de mod?les de conception de l?exp?rience utilisateur pour compl?ments Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code "projet de mod?les de conception de l?exp?rience utilisateur pour compl?ments Office") inclut des fichiers HTML, JavaScript et CSS que vous pouvez utiliser pour cr?er l?exp?rience utilisateur de votre compl?ment.   

Utiliser le projet de mod?les de conception d?exp?rience utilisateur aux fins suivantes :

* Appliquer des solutions ? des sc?narios client courants.
* Appliquer les meilleures pratiques en mati?re de conception.
* Incorporer les composants et styles d?[Office UI Fabric](https://dev.office.com/fabric#/get-started).
* Cr?er des compl?ments qui s?int?grent visuellement ? l?interface utilisateur d?Office par d?faut.  

## <a name="using-the-ux-design-patterns"></a>Utilisation des mod?les de conception UX

Vous pouvez utiliser le [Kit d'outils de conception de compl?ments Office](https://aka.ms/addins_toolkit) avec la [Kit d'outils de conception Fabric](https://aka.ms/fabric-toolkit) comme guide lorsque vous concevez votre propre compl?ment Office. Vous pouvez ?galement ajouter le [code source](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates) directement ? votre projet.

Pour utiliser les sp?cifications afin de cr?er une maquette de votre propre interface utilisateur du compl?ment, proc?dez comme suit :

1. T?l?chargez les fichiers de ressources de conception et commencez ? concevoir votre propre interface utilisateur :
    * [Kit d'outils de conception de compl?ments Office](https://aka.ms/addins_toolkit)
    * [Kit d'outils de conception Fabric](https://aka.ms/fabric-toolkit)

2. Pour obtenir des instructions, reportez-vous aux articles suivants :
    * Bonnes pratiques en mati?re de [conception de compl?ments Office](add-in-design.md)
    * [Kits d?outils Office UI Fabric](https://developer.microsoft.com/en-us/fabric#/resources)

> [!NOTE]
> Certains mod?les UX dans le kit d'outils de conception de compl?ments ne correspondent pas aux mod?les de conception UX d?taill?s ci-dessous. Nous pr?voyons de publier une documentation mise ? jour qui s'alignera sur le kit d'outils.

Pour ajouter le code source, proc?dez comme suit :

1. Clonez le [r?f?rentiel du projet de mod?les de conception de l?exp?rience utilisateur pour les compl?ments Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code "projet de mod?les de conception de l?exp?rience utilisateur pour les compl?ments Office").
2. Copiez le [dossier des composants](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/assets) ainsi que le dossier de code pour le mod?le individuel que vous choisissez dans votre projet de compl?ment.  
3. Incorporez le mod?le individuel ? votre compl?ment. Par exemple :
    - Modifiez l?emplacement source ou l?URL de commande de compl?ment dans le manifeste.
    - Utilisez le mod?le de conception d?exp?rience utilisateur en tant que mod?le pour d?autres pages.
    - Lien vers ou ? partir du mod?le de conception d?exp?rience utilisateur.

> [!NOTE]
> certaines sp?cifications de mod?le d?exp?rience utilisateur ne correspondent pas au code source. Nous mettons tout en ?uvre pour aligner toutes les ressources. Notez ?galement que certaines sp?cifications sont pr?sent?es comme archiv?es. Nous ?valuons la valeur de ces sp?cifications archiv?es sur la plateforme. Chaque mod?le vise ? repr?senter un mod?le unique et d?interaction. Les mod?les ne doivent pas se chevaucher et doivent ?tre bien diff?renci?s des composants Office Fabric UI.


## <a name="types-of-ux-design-patterns"></a>Types de mod?les de conception de l?exp?rience utilisateur
### <a name="generic-pages"></a>Pages g?n?riques

Les mod?les de page g?n?rique peuvent ?tre appliqu?s ? n?importe quelle page de votre compl?ment et n?ont pas d?usage particulier. L?un des mod?les de premi?re utilisation constitue un exemple de page ? usage sp?cifique. La liste suivante d?crit les pages g?n?riques disponibles :

* **Page d?accueil** : une page de compl?ment standard, par exemple la page sur laquelle un utilisateur est renvoy? apr?s une premi?re exp?rience d?utilisation ou un processus de connexion. 
    * En savoir plus sur les instructions relatives ? l?adoption du [langage de conception Office](add-in-design-language.md) dans votre compl?ment.
    * [Code de la page d?accueil](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/landing-page)
* **Image de marque dans la barre de marque** - La page d?accueil avec une image dans le pied de page qui repr?sente votre marque. 
    * [Sp?cification de la barre de marque](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/brand-bar.md)
    * [Code de la barre de marque](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/brand-bar)

<table>
 <tr><th>Accueil</th><th>Barre de marque</th></tr>
 <tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/landing-page"><img src="../images/landing-pages.png" alt="landing page" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/brand-bar"><img src="../images/word-brand-bar.png" alt="brand bar" style="width: 264px;"/></A></td></tr>
 </table>
 
### <a name="first-run-experience"></a>Premi?re exp?rience d?utilisation

Il s?agit de l?exp?rience v?cue par un utilisateur lorsqu?il ouvre votre compl?ment pour la premi?re fois. Les mod?les de mod?le de conception de premi?re utilisation suivants sont disponibles : 

* **?tapes de d?marrage** - Permet aux utilisateurs ayant une liste d??tapes ? suivre de commencer ? utiliser votre compl?ment. 
    * [?tapes de d?marrage d?une sp?cification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/fre_stepsToStart.pdf) (Ce mod?le de conception d?exp?rience utilisateur a ?t? archiv?. Comme nous ?valuons sa valeur, reportez-vous ? [Sp?cification sur la valeur de la premi?re ex?cution](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/value-placemat.md).)  
    * [Code des ?tapes de d?marrage](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/instruction-step)
* **Valeur** - Communique la proposition de valeur de votre compl?ment.
    * [Sp?cification de la valeur](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/value-placemat.md)
    * [Code de la valeur](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/value-placemat)
* **Vid?o** - Montre une vid?o aux utilisateurs avant qu?ils commencent ? utiliser votre compl?ment.
    * [Sp?cification de la vid?o](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/video-placemat.md)
    * [Code de la vid?o](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat)
* **Proc?dure pas ? pas** : explique aux utilisateurs une s?rie de fonctionnalit?s ou d?informations avant qu?ils commencent ? utiliser le compl?ment.
    * [Sp?cification de Carrousel](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/carousel.md) (Ce mod?le de conception d?exp?rience utilisateur a ?t? renomm? ? Carrousel ?. Les anciennes sp?cifications le d?signaient comme ? Panneau de pagination ?. Les ressources de code le d?signe comme ? Proc?dure pas ? pas pour la premi?re ex?cution ?. 
    * [Code de la proc?dure pas ? pas](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/walkthrough)

[AppSource](https://docs.microsoft.com/en-us/office/dev/store/use-the-seller-dashboard-to-submit-to-the-office-store) dispose d?un syst?me qui g?re les versions d??valuation d?un compl?ment, mais si vous souhaitez contr?ler l?interface utilisateur relative ? l?exp?rience d??valuation de votre compl?ment, utilisez les mod?les suivants :

* **Version d??valuation** - Explique aux utilisateurs comment utiliser la version d??valuation de votre compl?ment.
    * [Sp?cification d??valuation](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/fre_trialVersion.pdf) (Ce mod?le de conception d?exp?rience utilisateur a ?t? archiv?. Comme nous ?valuons sa valeur, reportez-vous ? ce PDF.)
    * [Code de la version d??valuation](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat)
* **Fonctionnalit? d??valuation** - Informe les utilisateurs que la fonctionnalit? qu?ils tentent d?utiliser n?est pas disponible dans la version d??valuation du compl?ment. Par ailleurs, si votre compl?ment est gratuit, mais qu?il comporte une fonctionnalit? qui n?cessite un abonnement, envisagez d?utiliser ce mod?le. Vous pouvez ?galement utiliser ce mod?le pour offrir une exp?rience avec une version ant?rieure apr?s qu?une p?riode d??valuation est termin?e.
    * [Sp?cification de la fonctionnalit? d??valuation](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/fre_trialFeature.pdf) (Ce mod?le de conception d?exp?rience utilisateur a ?t? archiv?. Comme nous ?valuons sa valeur, reportez-vous ? ce PDF.)
    * [Code de la fonctionnalit? d??valuation](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat-feature)

> [!IMPORTANT]
> Si vous d?cidez de g?rer votre propre version d??valuation et de ne pas utiliser AppSource pour g?rer la version d??valuation, assurez-vous que vous incluez la balise **Un autre achat peut ?tre requis** dans les notes de test du service Mon tableau de bord vendeur.

D?terminez s?il convient de montrer la vid?o sur la premi?re exp?rience d?utilisation une ou plusieurs fois (tout d?pend de son importance pour votre sc?nario). Par exemple, si les utilisateurs utilisent votre compl?ment r?guli?rement, ils peuvent oublier comment l?utiliser. Il peut ?tre utile de consulter la premi?re exp?rience d?utilisation plusieurs fois. 

 <table>
 <tr><th>?tapes de d?marrage</th><th>Valeur</th><th>Vid?o</th></tr>
 <tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/instruction-step"><img src="../images/instruction-steps.png" alt="instruction steps" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/value-placemat"><img src="../images/value-placemats.png" alt="value placemat" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat"><img src="../images/video-placemats.png" alt="video placemat" style="width: 250px;"/></A></td></tr>
 </table>

 <table>
 <tr><th>Premi?re page de la proc?dure pas ? pas</th><th>Version d??valuation</th><th>Fonctionnalit? d??valuation</th></tr>
 <tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/walkthrough"><img src="../images/walkthrough01.png" alt="walkthrough 1" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat"><img src="../images/trial-placemats.png" alt="trial placemat" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat-feature"><img src="../images/trial-placemats-feature.png" alt="trial placemat feature" style="width: 250px;"/></A></td></tr>
 </table> 

### <a name="navigation"></a>Navigation

Les utilisateurs doivent naviguer entre les diff?rentes pages de votre compl?ment. Les mod?les de navigation suivants indiquent diff?rentes options que vous pouvez utiliser afin d?organiser les pages et les commandes de votre compl?ment.

* **Bouton Page pr?c?dente et Page suivante** - Affiche un volet Office avec les boutons Page pr?c?dente et Page suivante. Utilisez ce mod?le pour vous assurer que les utilisateurs suivent une s?rie d??tapes ordonn?es.
    * [Sp?cification des boutons Page pr?c?dente et Page suivante](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/back-button.md)
    * [Code des boutons Page pr?c?dente et Page suivante](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/back-button) 
* **Navigation** - Affiche un menu, commun?ment appel? menu hamburger, avec les ?l?ments de menu de la page dans un volet Office. 
    * [Sp?cification de la navigation](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/contextual-menu.md)
    * [Code de la navigation](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation) 
* **Navigation ? l?aide de commandes** - Affiche le menu hamburger avec les boutons de commande (ou d?action) dans un volet Office. Utilisez ce mod?le lorsque vous voulez fournir des options de navigation et de commande ensemble. 
    * [Sp?cification de la navigation ? l?aide de commandes](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/command-bar.md)
    * [Code de la navigation ? l?aide de commandes](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation-commands)
* **Tableau crois? dynamique** - Affiche la navigation du tableau crois? dynamique dans un volet Office. Utilisez la navigation du tableau crois? dynamique pour permettre aux utilisateurs de naviguer entre les diff?rents contenus.
    * [Sp?cification du tableau crois? dynamique](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/pivot.md)
    * [Code du tableau crois? dynamique](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/pivot)
* **Barre d?onglets** - Affiche la navigation ? l?aide de boutons avec du texte et des ic?nes verticalement empil?s. Utiliser la barre d?onglets pour permettre la navigation ? l?aide des onglets avec des titres courts et explicites.
    * [Sp?cification de la barre d?onglets](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/tab-bar.md)
    * [Code de la barre d?onglets](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/tab-bar) 

<table>
<tr><th>Bouton Pr?c?dent</th><th>Navigation</th><th>Navigation ? l?aide de commandes</th></tr>
<tr>
    <td>
        <A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/back-button">
        <img src="../images/back-button.png" alt="back button" style="width: 250px;"/></A>
    </td>
    <td>
        <A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation">
        <img src="../images/navigation.png" alt="navigation" style="width: 250px;"/></A>
    </td>
    <td>
        <A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation-commands">
        <img src="../images/navigation-commands.png" alt="navigation with commands" style="width: 250px;"/></A>
    </td>
</tr>
 </table>

<table>
<tr><th>Pivot</th><th>Barre d?onglets</th></tr>
<tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/pivot">
<img src="../images/pivot.png" alt="pivot navigation" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/tab-bar">
<img src="../images/tab-bar.png" alt="tab bar" style="width: 250px;"/></A></td>
</tr>
 </table>

### <a name="notifications"></a>Notifications

Votre compl?ment peut avertir les utilisateurs d??v?nements, tels qu?une erreur, ou de l??tat d?avancement d?un ?l?ment de plusieurs fa?ons. Les mod?les de notification suivants sont disponibles : 

* **Bo?te de dialogue incorpor?e** - Affiche une bo?te de dialogue dans le volet des t?ches qui vous fournit des informations et, ?ventuellement, une exp?rience interactive, ? l?aide des boutons ou d?autres commandes. Pensez ? en utiliser une pour inviter un utilisateur ? confirmer une action. Utiliser le mod?le de bo?te de dialogue incorpor?e lorsque vous souhaitez conserver l?exp?rience utilisateur dans le volet Office.
    * [Sp?cification de la bo?te de dialogue incorpor?e](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/embedded-dialog.md)
    * [Code de la bo?te de dialogue incorpor?e](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/embedded-dialog)
* **Message incorpor?** - Indique l??chec, la r?ussite ou des informations, et peut appara?tre ? un emplacement sp?cifi? dans le volet Office. Par exemple, si un utilisateur entre une adresse de messagerie incorrecte dans une zone de texte, un message d?erreur appara?t juste en dessous de la zone de texte. 
    * [Sp?cification d?un message incorpor?](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/notification_inlineMessage.pdf) (Ce mod?le de conception d?exp?rience utilisateur a ?t? archiv?. Comme nous ?valuons sa valeur, reportez-vous ? ce PDF.)
    * [Code du message incorpor?](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/inline-message)
* **Banni?re de message** - Fournit des informations et, ?ventuellement, des instructions dans une banni?re qui peut ?tre r?duite ? une seule ligne, d?velopp?e en plusieurs lignes ou masqu?e. Utilisez des banni?res de message pour signaler une mise ? jour du service ou donner un conseil utile lorsque le compl?ment d?marre. 
    * [Sp?cification de banni?re de message](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/message_bar.pdf) (Ce mod?le de conception d?exp?rience utilisateur a ?t? archiv?. Comme nous ?valuons sa valeur, reportez-vous ? ce PDF.)
    * [Code de la banni?re de message](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/message-banner)
* **Barre de progression** - Indique la progression d?un processus long et synchrone, tel qu?une t?che de configuration qui doit ?tre termin?e pour que l?utilisateur puisse effectuer d?autres actions. Il s?agit d?une page distincte interstitielle qui met en ?vidence la marque du compl?ment. Utilisez une barre de progression quand le processus peut envoyer des notifications pour indiquer la progression de la t?che dans le compl?ment.
    * [Sp?cification de la barre de progression](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/progress-indicator.md)
    * [Code de la barre de progression](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/progress-bar)
* **Bouton fl?ch?** - Indique qu?un processus synchrone long est lanc?, mais ne fournit aucune indication sur son ?tat d?avancement. Il s?agit d?une page distincte interstitielle qui met en ?vidence la marque du compl?ment. Utilisez un bouton fl?ch? quand le compl?ment ne peut pas indiquer avec pr?cision la progression du processus. 
    * [Sp?cification du bouton fl?ch?](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/spinner.md)
    * [Code du bouton fl?ch?](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/spinner)
* **Annonce** - Fournit un bref message qui dispara?t au bout de quelques secondes. Comme il se peut que l?utilisateur ne voie pas le message, utilisez une annonce uniquement pour les informations non importantes. Utilisez une annonce pour informer les utilisateurs d?un ?v?nement dans un syst?me distant, tel que la r?ception d?un message ?lectronique.
    * [Sp?cification de l?annonce](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/toast.md)
    * [Code de l?annonce](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/toast)

 <table>
 <tr><th>Bo?te de dialogue incorpor?e</th><th>Message incorpor?</th><th>Banni?re de message</th></tr>
 <tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/embedded-dialog"><img src="../images/embedded-dialogs.png" alt="embedded dialog" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/inline-message"><img src="../images/inline-messages.png" alt="inline message" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/message-banner"><img src="../images/message-banners.png" alt="message banner" style="width: 250px;"/></A></td></tr>
 </table>

 <table>
 <tr><th>Barre de progression</th><th>Bouton fl?ch?</th><th>Annonce</th></tr>
 <tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/progress-bar"><img src="../images/progress-bars.png" alt="progress bar" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/spinner"><img src="../images/logo-spinner.png" alt="spinner" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/toast"><img src="../images/toast-header.png" alt="toast" style="width: 250px;"/></A></td></tr>
 </table>
 


### <a name="general-components"></a>Composants g?n?raux

Les ?l?ments suivants constituent des composants g?n?raux que vous pouvez utiliser avec vos compl?ments dans diff?rents sc?narios.  

#### <a name="client-dialog-boxes"></a>Bo?tes de dialogue client

Les bo?tes de dialogue client fournissent aux utilisateurs un autre moyen de travailler avec votre compl?ment en dehors d?un volet Office. Les mod?les de bo?te de dialogue suivants sont disponibles :

* **Bo?te de dialogue de rampe de type** - Affiche une bo?te de dialogue avec du contenu textuel. Utilisez la bo?te de dialogue de rampe de type pour transmettre des informations d?taill?es aux utilisateurs. 
    * Apprenez-en davantage sur la conception de [bo?tes de dialogue dans les compl?ments Office](dialog-boxes.md). Suivez ?galement nos recommandations concernant la [typographie dans les compl?ments Office](add-in-design-language.md#typography).
    * [Code de la bo?te de dialogue de rampe de type](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/typeramp)
* **Bo?te de dialogue d?alerte** - Affiche un message d?alerte avec des informations importantes, comme les erreurs ou les notifications, aux utilisateurs.  
    * [Sp?cification de bo?te de dialogue d?alerte](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/notification_alert.pdf) (Ce mod?le de conception d?exp?rience utilisateur a ?t? archiv?. Comme nous ?valuons sa valeur, reportez-vous ? ce PDF.)
    * [Code de la bo?te de dialogue d?alerte](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/alert)
* **Bo?te de dialogue de navigation** - Affiche une bo?te de dialogue comportant la navigation. Utilisez la bo?te de dialogue de navigation pour permettre aux utilisateurs de naviguer entre les diff?rents contenus. 
    * Apprenez-en davantage sur la conception de [bo?tes de dialogue dans des compl?ments Office](dialog-boxes.md). D?couvrez ?galement comment utiliser les [composants de tableau crois? dynamique Office UI Fabric dans les compl?ments Office](pivot.md).
    * [Code de la bo?te de dialogue de navigation](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation)

<table>
 <tr><th>Bo?te de dialogue de rampe de type</th><th>Bo?te de dialogue d?alerte</th></tr>
<tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/typeramp"><img src="../images/typeramp-dialog.png" alt="typeramp dialog" width="400"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/alert"><img src="../images/alert-dialog.png" alt="alert dialog" width="400"/></A></td>
</tr></tr>
 </table>
 
 <table>
 <tr><th>Bo?te de dialogue de navigation</th></tr>
<tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation"><img src="../images/navigation-dialog.png" alt="navigation dialog" width="450"/></A></td></tr>
</tr>
 </table>


#### <a name="feedback-and-ratings"></a>?valuations et commentaires

Pour am?liorer la visibilit? et l?adoption de votre compl?ment, il est utile de fournir aux utilisateurs la possibilit? de noter et de commenter votre compl?ment dans AppSource. Ce mod?le comporte deux m?thodes pour effectuer des commentaires et des ?valuations dans le compl?ment :

- Commentaires initi?s par l?utilisateur - Un utilisateur choisit d?envoyer des commentaires ? l?aide du menu de navigation (par exemple, en utilisant le lien **Envoyer des commentaires**) ou d?une ic?ne dans le pied de page.
- Commentaires initi?s par le syst?me - Une fois le compl?ment ex?cut? trois fois, l?utilisateur est invit? ? fournir un commentaire, via une banni?re de message.

Les deux m?thodes ouvrent une bo?te de dialogue qui contient la page AppSource pour le compl?ment.

* [Sp?cification des ?valuations et des commentaires](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/notification_feedback.pdf) (Ce mod?le de conception d?exp?rience utilisateur a ?t? archiv?. Comme nous ?valuons sa valeur, reportez-vous ? ce PDF.)
* [Code des ?valuations et commentaires](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/feedback/office-store)

> [!IMPORTANT]
> Ce mod?le pointe actuellement vers la page d?accueil d?AppSource. Veillez ? mettre ? jour l?URL avec l?URL de la page de votre compl?ment dans AppSource.


 <table>
 <tr><th>?valuations et commentaires</th></tr>
<tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/feedback/office-store"><img src="../images/feedback-rating.png" alt="Feedback and Ratings" style="width: 350px;"/></A></td></tr>
</tr>
 </table>

#### <a name="settings-and-privacy"></a>Param?tres et confidentialit?

Les compl?ments peuvent n?cessiter une page des param?tres afin que les utilisateurs puissent configurer les param?tres qui contr?lent le comportement du compl?ment. Vous pouvez ?galement fournir aux utilisateurs les politiques de confidentialit? auxquelles votre compl?ment adh?re. 

* **Param?tres** - Affiche un volet Office avec des composants de configuration contr?lant le comportement du compl?ment. Une page des param?tres fournit des options que l?utilisateur peut choisir.
    * [Sp?cification des param?tres](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/settings.md)
    * [Code des param?tres](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings)
* **Politique de confidentialit?** - Affiche un volet Office contenant des informations importantes sur les politiques de confidentialit?. 
    * [Sp?cification de la politique de confidentialit?](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/general_multiSection.pdf) (Ce mod?le de conception d?exp?rience utilisateur a ?t? archiv?. Comme nous ?valuons sa valeur, reportez-vous ? ce PDF.)
    * [Code de la politique de confidentialit?](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings)

<table>
 <tr><th>Param?tres</th><th>Politique de confidentialit?</th></tr>
<tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings"><img src="../images/settings.png" alt="settings" style="width: 300px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings"><img src="../images/privacy-policy.png" alt="privacy" style="width: 264px;"/></A></td>
</tr></tr>
 </table>

## <a name="see-also"></a>Voir aussi

* [Meilleures pratiques en mati?re de d?veloppement de compl?ments Office](../concepts/add-in-development-best-practices.md)
* [Office UI Fabric](http://dev.office.com/fabric/)
