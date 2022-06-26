---
title: glossaire des termes des compléments Office
description: Glossaire des termes couramment utilisés dans la documentation des compléments Office.
ms.date: 06/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: 002c61cf482da75a5fa2bef0219990ffc9b04034
ms.sourcegitcommit: d8fbe472b35c758753e5d2e4b905a5973e4f7b52
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/25/2022
ms.locfileid: "66229644"
---
# <a name="office-add-ins-glossary"></a>glossaire des compléments Office

Il s’agit d’un glossaire des termes couramment utilisés dans la documentation des compléments Office.

## <a name="add-in"></a>macro complémentaire

Office compléments sont des applications web qui étendent Office applications. Ces applications web ajoutent de nouvelles fonctionnalités à l’application Office, telles que l’ajout de données externes, l’automatisation des processus ou l’incorporation d’objets interactifs dans Office documents.

Office compléments diffèrent des compléments VBA, COM et VSTO, car ils offrent une prise en charge multiplateforme (généralement web, Windows, Mac et iPad) et sont basés sur des technologies web standard (HTML, CSS et JavaScript). Le langage de programmation principal d’un complément Office est JavaScript ou TypeScript.

## <a name="add-in-commands"></a>commandes de complément

**Les commandes de complément sont des** éléments d’interface utilisateur, tels que des boutons et des menus, qui étendent l’interface utilisateur Office pour votre complément. Lorsque les utilisateurs sélectionnent un élément de commande de complément, ils lancent des actions telles que l’exécution de code JavaScript ou l’affichage du complément dans un volet Office. Les commandes de complément permettent à votre complément de ressembler à une partie de Office, ce qui donne aux utilisateurs plus de confiance dans votre complément. Pour plus d’informations[, consultez les commandes de complément pour Excel, PowerPoint et Word](../design/add-in-commands.md) et les [commandes de complément pour Outlook](../outlook/add-in-commands-for-outlook.md).

Voir aussi : [ruban, bouton du ruban](#ribbon-ribbon-button).

## <a name="application"></a>application

**L’application** fait référence à une application Office. Les applications Office qui prennent en charge les compléments Office sont Excel, OneNote, Outlook, PowerPoint, Project et Word.

Voir aussi : [client](#client), [hôte](#host), [application Office, client Office](#office-application-office-client).

## <a name="application-specific-api"></a>API spécifique à l’application

Les API spécifiques à l’application fournissent des objets fortement typés qui interagissent avec des objets natifs d’une application Office spécifique. Par exemple, vous appelez les API JavaScript Excel pour accéder à des feuilles de calcul, des plages, des tableaux, des graphiques, etc. Des API spécifiques à l’application sont actuellement disponibles pour Excel, OneNote, PowerPoint, Visio et Word. Pour en savoir plus, consultez le [modèle d’API spécifique à l’application](../develop/application-specific-api-model.md) .

Voir aussi : [API courante](#common-api).

## <a name="client"></a>Client

**Le client** fait généralement référence à une application Office. Les applications Office ou les clients qui prennent en charge les compléments Office sont Excel, OneNote, Outlook, PowerPoint, Project et Word.

Voir aussi : [application](#application), [hôte](#host), [application Office, client Office](#office-application-office-client).

## <a name="common-api"></a>API communes

Les API courantes sont utilisées pour accéder à des fonctionnalités telles que l’interface utilisateur, les boîtes de dialogue et les paramètres clients qui sont communs à plusieurs applications Office. Ce modèle d’API utilise des[rappels](https://developer.mozilla.org/docs/Glossary/Callback_function), qui vous permettent de spécifier une seule opération dans chaque demande envoyée à l’application Office.

Les API courantes ont été introduites avec Office 2013 et sont utilisées pour interagir avec Office 2013 ou version ultérieure. Certaines API courantes sont des API héritées du début des années 2010. Excel, PowerPoint et Word ont tous des fonctionnalités d’API communes, mais la plupart de ces fonctionnalités ont été remplacées ou remplacées par le modèle d’API spécifique à l’application. Les API spécifiques à l’application sont recommandées lorsque cela est possible.

D’autres API courantes, telles que les API communes liées à Outlook, à l’interface utilisateur et à l’authentification, sont les API modernes et préférées à ces fins. Pour plus d’informations sur le modèle objet d’API common, consultez [Common JavaScript API object model](../develop/office-javascript-api-object-model.md).

Voir aussi : [API spécifique à l’application](#application-specific-api).

## <a name="content-add-in"></a>complément de contenu

**Les compléments de contenu sont des** vues web ou des vues de navigateur web qui sont incorporées directement dans des documents Excel, OneNote ou PowerPoint. Les compléments de contenu permettent aux utilisateurs d’accéder aux contrôles d’interface qui exécutent le code pour modifier des documents ou afficher des données d’une source de données. Utilisez les compléments de contenu lorsque vous souhaitez incorporer des fonctionnalités directement dans le document. Pour en savoir plus, consultez [Les compléments Office de contenu](../design/content-add-ins.md).

Voir aussi : [webview](#webview).

## <a name="content-delivery-network-cdn"></a>réseau de distribution de contenu (CDN)

Un **réseau de distribution de contenu** ou **CDN** est un réseau distribué de serveurs et de centres de données. Il offre généralement une disponibilité et des performances des ressources plus élevées par rapport à un seul serveur ou centre de données.

## <a name="contoso"></a>Contoso

**Contoso** Ltd. (également appelée Contoso et Contoso University) est une société fictive utilisée par Microsoft comme exemple de société et de domaine.

## <a name="custom-function"></a>fonction personnalisée

Une **fonction personnalisée** est une fonction définie par l’utilisateur qui est empaqueté avec un complément Excel. Les fonctions personnalisées permettent aux développeurs d’ajouter de nouvelles fonctions, au-delà des fonctionnalités de Excel classiques, en définissant ces fonctions dans JavaScript dans le cadre d’un complément. Les utilisateurs de Excel peuvent accéder aux fonctions personnalisées comme n’importe quelle fonction native dans Excel. Pour en savoir plus[, consultez Créer des fonctions personnalisées dans Excel](../excel/custom-functions-overview.md).

## <a name="custom-functions-runtime"></a>runtime de fonctions personnalisées

Un **runtime de fonctions personnalisées** est un runtime JavaScript uniquement qui exécute uniquement des fonctions personnalisées. Il n’a pas d’interface utilisateur et ne peut pas interagir avec Office.js API. Si votre complément a uniquement des fonctions personnalisées, il s’agit d’un bon runtime léger à utiliser. Si vos fonctions personnalisées doivent interagir avec le volet Office ou les API Office.js, configurez un runtime JavaScript partagé. Pour plus d’information, consultez [Configurer votre complément Office pour utiliser un runtime JavaScript partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

Voir aussi : [Runtime JavaScript](#javascript-runtime), [runtime JavaScript partagé, runtime partagé](#shared-javascript-runtime-shared-runtime).

## <a name="host"></a>host

**L’hôte** fait généralement référence à une application Office. Les applications ou hôtes Office qui prennent en charge les compléments Office sont Excel, OneNote, Outlook, PowerPoint, Project et Word.

Voir aussi : [application](#application), [client](#client), [application Office, client Office](#office-application-office-client).

## <a name="javascript-runtime"></a>Runtime JavaScript

Le **runtime JavaScript** est l’environnement hôte du navigateur dans lequel le complément s’exécute. Dans Office sur Windows et Office sur Mac, le runtime JavaScript est un contrôle de navigateur incorporé (ou webview) tel qu’Internet Explorer, Edge Legacy, Edge WebView2 ou Safari. Différentes parties d’une exécution de complément dans des runtimes JavaScript distincts. Par exemple, les commandes de complément, les fonctions personnalisées et le code du volet Office utilisent généralement des runtimes JavaScript distincts, sauf si vous configurez un runtime JavaScript partagé. Pour plus d’informations, consultez [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md).

Voir aussi : [runtime de fonctions personnalisées](#custom-functions-runtime), [runtime JavaScript partagé, runtime partagé](#shared-javascript-runtime-shared-runtime), [webview](#webview).

## <a name="office-application-office-client"></a>Office application, client Office

**Office client** fait référence à une application Office. Les applications Office ou les clients qui prennent en charge les compléments Office sont Excel, OneNote, Outlook, PowerPoint, Project et Word.

Voir aussi : [application](#application), [client](#client), [hôte](#host).

## <a name="platform"></a>platform

Une **plateforme** fait généralement référence au système d’exploitation exécutant l’application Office. Les plateformes qui prennent en charge les compléments Office incluent les navigateurs Windows, Mac, iPad et web.

## <a name="quick-start"></a>démarrage rapide

Un **démarrage rapide** est une description générale des compétences clés et des connaissances requises pour le fonctionnement de base d’un programme particulier. Dans la documentation Office Compléments, un démarrage rapide est une introduction au développement d’un complément pour une application particulière, telle que Outlook. Un démarrage rapide contient une série d’étapes qu’un développeur de compléments peut effectuer en environ 5 minutes, ce qui entraîne un complément fonctionnel et un environnement de développement fonctionnel.

Voir aussi : [didacticiel](#tutorial).

## <a name="requirement-set"></a>ensemble de conditions requises

[!include[Requirement set note](../includes/office-js-requirement-sets.md)]

## <a name="ribbon-ribbon-button"></a>ruban, bouton du ruban

Un **ruban** est une barre de commandes qui organise les fonctionnalités d’une application en une série d’onglets ou de boutons en haut d’une fenêtre. Un **bouton de ruban** est l’un des boutons de cette série. Pour plus d’informations, consultez [Afficher ou masquer le ruban dans Office](https://support.microsoft.com/office/d946b26e-0c8c-402d-a0f7-c6efa296b527#ID0EBBD=Newer_Versions).

## <a name="runtime"></a>Duree

Voir : [Runtime JavaScript](#javascript-runtime).

## <a name="shared-javascript-runtime-shared-runtime"></a>runtime JavaScript partagé, runtime partagé

Un **runtime JavaScript partagé**, ou **runtime partagé**, permet à tout le code de votre complément, y compris le volet Office, les commandes de complément et les fonctions personnalisées, de s’exécuter dans le même runtime JavaScript et de continuer à s’exécuter même lorsque le volet Office est fermé. Pour en savoir plus, consultez [Configurer votre complément Office pour utiliser un runtime JavaScript partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md) et [Astuces pour utiliser le runtime JavaScript partagé dans votre complément Office](https://devblogs.microsoft.com/microsoft365dev/tips-for-using-the-shared-javascript-runtime-in-your-office-add-in%e2%80%af/).

Voir aussi : [runtime de fonctions personnalisées](#custom-functions-runtime), [runtime JavaScript](#javascript-runtime).

## <a name="task-pane"></a>volet Office

Les volets Office sont des surfaces d’interface ou des vues web qui apparaissent généralement sur le côté droit de la fenêtre dans Excel, Outlook, PowerPoint et Word. Les volets des tâches permettent aux utilisateurs d’accéder aux contrôles d’interface qui exécutent le code pour modifier des documents ou des e-mails, ou afficher des données d’une source de données. Utilisez les volets Office lorsque vous n’avez pas besoin ou ne pouvez pas incorporer de fonctionnalités directement dans le document. Pour en savoir plus, consultez [les volets Office dans Office compléments](../design/task-pane-add-ins.md).

Voir aussi : [webview](#webview).

## <a name="tutorial"></a>Tutoriel

Un **didacticiel** est une aide pédagogique conçue pour aider les gens à apprendre à utiliser un produit ou une procédure. Dans le contexte des compléments Office, un didacticiel guide un développeur de compléments tout au long du processus de développement complet d’un complément pour une application particulière, telle que Excel. Cela implique de suivre 20 étapes ou plus et représente un investissement de temps plus important qu’un [démarrage rapide](#quick-start).

Voir aussi : [démarrage rapide](#quick-start).

## <a name="custom-functions-only-add-in"></a>complément de fonctions personnalisées uniquement

Complément qui contient une fonction personnalisée, mais aucune interface utilisateur telle qu’un volet Office. Les fonctions personnalisées de ce type de complément s’exécutent dans un runtime JavaScript uniquement. Une fonction personnalisée qui inclut une interface utilisateur peut utiliser un runtime partagé ou une combinaison d’un runtime JavaScript uniquement et d’un runtime html. Nous vous recommandons d’utiliser un runtime partagé si vous disposez d’une interface utilisateur. 

Voir aussi : [fonction personnalisée](#custom-function), [runtime de fonctions personnalisées](#custom-functions-runtime).

## <a name="web-add-in"></a>complément web

**Le complément web** est un terme hérité pour un complément Office. Ce terme peut être utilisé lorsque la documentation Microsoft 365 doit distinguer les compléments Office modernes des autres types de compléments tels que VBA, COM ou VSTO.

Voir aussi : [complément](#add-in).

## <a name="webview"></a>Webview

Une **vue web** est un élément ou une vue qui affiche du contenu web à l’intérieur d’une application. Les compléments de contenu et les volets Office contiennent tous deux des navigateurs web incorporés et sont des exemples de vues web dans Office compléments.

Voir aussi : [complément de contenu](#content-add-in), [volet Office](#task-pane).

## <a name="xll"></a>XLL

Un complément **XLL** est un fichier de complément Excel qui fournit des fonctions définies par l’utilisateur et possède l’extension de fichier **.xll**. Un fichier XLL est un type de fichier de bibliothèque de liens dynamiques (DLL) qui ne peut être ouvert que par Excel. Les fichiers de complément XLL doivent être écrits en C ou C++. Les fonctions personnalisées sont l’équivalent moderne des fonctions XLL définies par l’utilisateur. Les fonctions personnalisées offrent une prise en charge sur plusieurs plateformes et sont rétrocompatibles avec les fichiers XLL. Pour plus d’informations, consultez [Étendre des fonctions personnalisées avec des fonctions XLL définies par l’utilisateur](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf) .

Voir aussi : [fonction personnalisée](#custom-function).

## <a name="yeoman-generator-yo-office"></a>Générateur Yeoman, yo office

Le [générateur Yeoman pour les compléments Office](../develop/yeoman-generator-overview.md) utilise l’outil open source [Yeoman](https://github.com/yeoman/yo) pour générer un complément Office via la ligne de commande. `yo office`est la commande qui exécute le générateur Yeoman pour Office compléments. Les Office compléments démarrent rapidement et les didacticiels utilisent le générateur Yeoman.

## <a name="see-also"></a>Voir aussi

- [Ressources supplémentaires sur les compléments Office](resources-links-help.md)