---
title: Vue d?ensemble de la plateforme des compl?ments pour Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: f0f20371eee759a449773effaff1ce365e32bf48
ms.sourcegitcommit: 17f60431644b448a4816913039aaebfa328f9b0a
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/25/2018
---
# <a name="office-add-ins-platform-overview"></a>Vue d?ensemble de la plateforme de compl?ments pour Office

La plateforme des compl?ments Office permet de cr?er des solutions qui ?tendent des applications Office et interagissent avec du contenu dans des documents Office. Les compl?ments Office vous permettent d?utiliser des technologies web que vous connaissez, telles que le code HTML, CSS et JavaScript, pour ?tendre Word, Excel, PowerPoint, OneNote, Project et Outlook, et interagir avec ces programmes. Votre solution peut ?tre ex?cut?e dans Office sur plusieurs plateformes, notamment Office pour Windows, Office Online, Office pour Mac et Office pour iPad.

Les compl?ments Office offrent presque les m?mes possibilit?s qu?une page web dans un navigateur. Vous pouvez utiliser la plateforme des compl?ments Office pour :

-  **Ajout de nouvelles fonctionnalit?s ? des clients Office :** vous pouvez importer des donn?es externes dans Office, automatiser des documents Office, exposer des fonctionnalit?s tierces dans des clients Office et bien plus encore. Par exemple, vous pouvez utiliser l?API Microsoft Graph pour ?tablir une connexion vers des donn?es qui am?liorent la productivit?. 
    
-  **Cr?er de nouveaux objets interactifs et enrichis qui peuvent ?tre incorpor?s dans des documents Office :** vous pouvez incorporer des cartes, des graphiques et des visualisations interactives que les utilisateurs peuvent ajouter ? leurs feuilles de calcul Excel et pr?sentations PowerPoint. 
    
## <a name="how-are-office-add-ins-different-than-com-and-vsto-add-ins"></a>En quoi les compl?ments Office sont-ils diff?rents des compl?ments COM et VSTO ? 

Les compl?ments COM ou VSTO sont des solutions d?int?gration ? Office ant?rieures qui s?ex?cutent uniquement sur Office pour Windows. Contrairement aux compl?ments COM, les compl?ments Office n?incluent pas de code ex?cut? sur l?appareil de l?utilisateur ou sur le client Office. Pour un compl?ment Office, l?application h?te, par exemple Excel, lit le manifeste du compl?ment et ins?re les commandes de menu et les boutons de ruban personnalis?s du compl?ment dans l?interface utilisateur. Lorsque cela est n?cessaire, elle charge le code JavaScript et HTML du compl?ment, qui est ex?cut? dans le contexte d?un navigateur dans un bac ? sable (sandbox). 

Les compl?ments Office offrent les avantages suivants par rapport aux compl?ments cr??s ? l?aide de VBA, COM ou VSTO : 

- Prise en charge sur plusieurs plateformes. Les compl?ments Office s?ex?cutent dans Office pour Windows, Mac, iOS et Office Online. 

- Authentification unique (SSO) : les compl?ments Office s?int?grent facilement ? des comptes d?utilisateurs Office 365. 

- D?ploiement et distribution centralis?s. Les administrateurs peuvent d?ployer des compl?ments Office de fa?on centralis?e dans une organisation. 

- Acc?s facile via AppSource. Vous pouvez mettre votre solution ? disposition d?un large public en l?envoyant ? AppSource. 

- S?appuie sur des technologies web standard. Vous pouvez utiliser n?importe quelle biblioth?que pour cr?er des compl?ments Office. 

## <a name="components-of-an-office-add-in"></a>Composants d?un compl?ment Office 

Un compl?ment Office inclut deux composants de base : un fichier manifeste XML et votre propre application web. Le manifeste d?finit diff?rents param?tres, y compris la fa?on dont votre compl?ment s?int?gre avec les clients Office. Votre application web doit ?tre h?berg?e sur un serveur web ou un service d?h?bergement web, tel que Microsoft Azure.

*Figure 1. Manifeste + page web = compl?ment Office*

![Manifeste + page web = compl?ment Office](../images/dk2-agave-overview-01.png)

### <a name="manifest"></a>Manifeste 

Le manifeste est un fichier XML qui sp?cifie les param?tres et les fonctionnalit?s du compl?ment, notamment : 

- Le nom d?affichage, la description, l?ID, la version et les param?tres r?gionaux par d?faut du compl?ment. 

- La fa?on dont le compl?ment s?int?gre ? Office.  

- Le niveau d?autorisation et les conditions d?acc?s aux donn?es pour le compl?ment. 

### <a name="web-app"></a>Application web 

Le compl?ment Office le plus simple est compos? d?une page HTML statique qui est affich?e dans une application Office, mais qui n?interagit pas avec le document Office ou une autre ressource Internet. Toutefois, pour cr?er un compl?ment qui interagit avec des documents Office ou permet ? l?utilisateur d?interagir avec les ressources en ligne ? partir d?une application h?te Office, vous pouvez utiliser n?importe quelle technologie, aussi bien c?t? client que serveur, prise en charge par votre fournisseur d?h?bergement (par exemple, ASP.NET, PHP ou Node.js). Pour interagir avec des clients et des documents Office, vous pouvez utiliser les API JavaScript Office.js. 

*Figure 2. Composants d?un compl?ment Office Hello World*

![Composants d?un compl?ment Hello World](../images/dk2-agave-overview-07.png)

## <a name="extending-and-interacting-with-office-clients"></a>Extension des clients Office et interaction avec ces clients 

Les compl?ments Office offrent les possibilit?s suivantes dans une application Office h?te : 

-  ?tendre les fonctionnalit?s (toutes les applications Office) 

-  Cr?er de nouveaux objets (Excel ou PowerPoint) 
 
### <a name="extend-office-functionality"></a>?tendre les fonctionnalit?s d?Office 

Vous pouvez ajouter de nouvelles fonctionnalit?s aux applications Office via les ?l?ments d?interface suivants :  

-  Commandes de menu et boutons de ruban personnalis?es (collectivement appel?s ? commandes de compl?ment ?) 

-  Volets Office ? ins?rer 

Les ?l?ments d?interface personnalis?s et les volets Office sont d?finis dans le manifeste du compl?ment.  

#### <a name="custom-buttons-and-menu-commands"></a>Commandes de menu et boutons personnalis?s  

Vous pouvez ajouter des ?l?ments de menu et des boutons de ruban personnalis? au ruban d?Office pour bureau Windows et Office Online. Les utilisateurs peuvent ainsi acc?der ? votre compl?ment directement ? partir de leur application Office. Les boutons de commande peuvent lancer diff?rentes actions, par exemple afficher un volet Office comportant du contenu HTML personnalis? ou ex?cuter une fonction JavaScript.  

*Figure 3. Commandes de compl?ment en cours d?ex?cution dans Excel (version de bureau)*

![Commandes de menu et boutons personnalis?s](../images/add-in-commands-overview.png)

#### <a name="task-panes"></a>Volets Office  

Vous pouvez utiliser des volets Office en plus des commandes de compl?ment pour permettre aux utilisateurs d?interagir avec votre solution. Les clients qui ne prennent pas en charge les commandes de compl?ment (Office 2013 et Office pour iPad) ex?cutent votre compl?ment sous la forme d?un volet Office. Les utilisateurs lancent les compl?ments de volet Office via le bouton **Mes compl?ments** situ? sous l?onglet **Insertion**. 

*Figure 4. Volet Office*

![Volet de t?ches](../images/task-pane-overview.jpg)

### <a name="extend-outlook-functionality"></a>Extension des fonctionnalit?s Outlook 

Les compl?ments Outlook peuvent d?velopper le ruban Office et s?afficher en regard d?un ?l?ment Outlook quand vous le visualisez ou le composez. Ils fonctionnent avec un message ?lectronique, une demande de r?union, une r?ponse ? une demande de r?union, une annulation de r?union ou un rendez-vous quand l?utilisateur visualise un ?l?ment re?u, r?pond ? un ?l?ment ou en cr?e un. 

Les compl?ments Outlook peuvent acc?der ? des informations contextuelles ? partir de l??l?ment, telles qu?une adresse ou un ID de suivi, et utiliser ces donn?es pour acc?der ? d?autres informations sur le serveur ou provenant de services web pour cr?er des exp?riences utilisateur attrayantes. Dans la plupart des cas, un compl?ment Outlook peut ?tre ex?cut? sans modification sur les diff?rentes applications h?te prise en charge, notamment Outlook, Outlook pour Mac, Outlook Web App et Outlook Web App pour appareils, afin d?offrir une exp?rience homog?ne sur le bureau, en ligne, sur les tablettes et sur les appareils mobiles. 

Pour acc?der ? une vue d?ensemble des compl?ments Outlook, reportez-vous ? la rubrique [Pr?sentation des compl?ments Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/). 

### <a name="create-new-objects-in-office-documents"></a>Cr?ation d?objets dans des documents Office 

Vous pouvez incorporer des objets web, appel?s compl?ments de contenu, dans des documents Excel et PowerPoint. Ces compl?ments de contenu vous permettent d?int?grer des visualisations de donn?es web enrichies, du contenu multim?dia (comme un lecteur vid?o YouTube ou une galerie d?images) et d?autres types de contenu externe.

*Figure 5. Compl?ment de contenu*

![compl?ment de contenu](../images/dk2-agave-overview-05.png)

## <a name="office-javascript-apis"></a>API JavaScript pour Office 

Les API JavaScript Office sont compos?es d?objets et de membres permettant de cr?er des compl?ments et d?interagir avec le contenu Office et les services web. Il existe un mod?le objet commun que se partagent Excel, Outlook, Word, PowerPoint, OneNote et Project. Il existe ?galement des mod?les objet plus complets et propres ? l?h?te pour Excel et Word. Ces API permettent d?acc?der ? des objets connus tels que des paragraphes et des classeurs, ce qui facilite la cr?ation de compl?ment pour un h?te sp?cifique.  

## <a name="next-steps"></a>?tapes suivantes 

Pour en savoir plus sur la cr?ation de votre compl?ment Office, essayez notre [D?marrage rapide en 5 minutes](https://docs.microsoft.com/en-us/office/dev/add-ins/). Vous pouvez commencer ? cr?er des compl?ments imm?diatement ? l'aide de Visual Studio ou de tout autre ?diteur. 

Pour commencer ? concevoir des solutions offrant des exp?riences utilisateur efficaces et attrayantes, consultez les [instructions de conception](../design/add-in-design.md) et les [meilleures pratiques](../concepts/add-in-development-best-practices.md) pour les compl?ments Office.    
   
## <a name="see-also"></a>Voir aussi

- [Exemples de compl?ments Office](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples)
- [Pr?sentation de l?API JavaScript pour Office](../develop/understanding-the-javascript-api-for-office.md)
- [Disponibilit? des compl?ments Office sur les plateformes et les h?tes](../overview/office-add-in-availability.md)


    
